"""
DocxAnnotator
=============
Annotates a DOCX file with format errors detected by the paper-format-modifier
checker.  Accepts the same errors_list structure that the frontend produces and
writes visible markers directly into a copy of the original document:

  * Character-level errors  → red + bold text
  * Ghost references [n]    → yellow highlight + Word comment
  * Layout/structure errors → Word comment balloon (no text change)

Two-phase approach
------------------
1. Open the DOCX with python-docx, modify paragraph lxml elements in-memory
   (colours, highlights, w:commentRangeStart/End/Reference markers).
2. Save to bytes, then post-process the ZIP:
   - inject  word/comments.xml
   - patch   [Content_Types].xml
   - patch   word/_rels/document.xml.rels
"""

from __future__ import annotations

import io
import re
import zipfile
from datetime import datetime
from lxml import etree
from docx import Document

# ---------------------------------------------------------------------------
# OOXML constants
# ---------------------------------------------------------------------------
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"
NSMAP = {"w": W_NS}

COMMENTS_CT = (
    "application/vnd.openxmlformats-officedocument"
    ".wordprocessingml.comments+xml"
)


def _w(tag: str) -> str:
    """Return a Clark-notation tag in the W namespace."""
    return f"{{{W_NS}}}{tag}"


# ---------------------------------------------------------------------------
# DocxAnnotator
# ---------------------------------------------------------------------------
class DocxAnnotator:
    """Annotate a DOCX file with a list of format errors."""

    # ------------------------------------------------------------------
    # Public entry point
    # ------------------------------------------------------------------

    def annotate_document(
        self,
        input_source,
        errors_list: list[dict],
        output_path: str | None = None,
    ) -> bytes:
        """
        Annotate a DOCX and return the result as bytes.

        Parameters
        ----------
        input_source : bytes | BytesIO | str (file path)
        errors_list  : list of error dicts from the frontend checker.
            Minimum keys: type, level, title, desc, suggestion, location.
            Optional: location as dict with paragraph_idx, start, end.
        output_path  : if given, also write the result to this path.
        """
        if isinstance(input_source, (bytes, bytearray)):
            buf: io.IOBase = io.BytesIO(bytes(input_source))
        elif isinstance(input_source, io.IOBase):
            buf = input_source
        else:
            buf = open(input_source, "rb")  # type: ignore[assignment]

        doc = Document(buf)
        self._next_id: int = 0
        self._comment_records: list[dict] = []

        paragraphs = doc.paragraphs

        for error in errors_list:
            try:
                self._process_error(doc, paragraphs, error)
            except Exception:
                # Never let a single bad error crash the whole annotation
                pass

        # Serialize (python-docx preserves our lxml modifications)
        out = io.BytesIO()
        doc.save(out)
        out.seek(0)
        raw = out.read()

        # Post-process ZIP to inject comments.xml
        if self._comment_records:
            raw = self._inject_comments_into_zip(raw)

        if output_path:
            with open(output_path, "wb") as f:
                f.write(raw)

        return raw

    # ------------------------------------------------------------------
    # Error routing
    # ------------------------------------------------------------------

    def _process_error(self, doc, paragraphs: list, error: dict) -> None:
        etype = error.get("type", "")
        level = error.get("level", "L3")
        title = error.get("title", "")
        desc  = error.get("desc",  "")
        sugg  = error.get("suggestion", "")
        loc   = error.get("location", "")

        # Case A: explicit paragraph_idx supplied
        if isinstance(loc, dict) and loc.get("paragraph_idx") is not None:
            pidx = int(loc["paragraph_idx"])
            if pidx < len(paragraphs):
                para = paragraphs[pidx]
                start = loc.get("start")
                end   = loc.get("end")
                if etype in ("font_size", "font_family", "char_spacing", "typo"):
                    self._add_red_text_marker(para, start, end)
                comment = self._fmt_comment(level, title, etype, loc, desc, sugg)
                self._add_comment(doc, para, comment)
            return

        # Case B: ghost references
        if etype == "引用完整性" and "幽灵引用" in title:
            self._handle_ghost_refs(doc, paragraphs, error)
            return

        # Case C: punctuation / mixed script
        if etype == "中英文混排":
            self._handle_punct_error(doc, paragraphs, error)
            return

        # Case D: generic layout / structure comment
        self._handle_generic(doc, paragraphs, error)

    # ------------------------------------------------------------------
    # Annotation strategies
    # ------------------------------------------------------------------

    def _handle_ghost_refs(self, doc, paragraphs: list, error: dict) -> None:
        """Yellow-highlight every [n] ghost citation and attach a comment."""
        level = error.get("level", "L1")
        title = error.get("title", "")
        desc  = error.get("desc",  "")
        sugg  = error.get("suggestion", "")

        # Extract numbers from strings like "正文引用了 [1, 5, 12]"
        ghost_nums = re.findall(r"\[(\d+)\]", desc)
        if not ghost_nums:
            # Also try bare comma-separated list after last bracket
            m = re.search(r"\[([0-9,\s]+)\]", desc)
            if m:
                ghost_nums = [n.strip() for n in m.group(1).split(",")]

        annotated: set[str] = set()
        for para in paragraphs:
            para_text = para.text
            for gn in ghost_nums:
                pattern = f"[{gn}]"
                if pattern in para_text and gn not in annotated:
                    self._add_yellow_highlight_to_para(para, pattern)
                    comment = (
                        f"[格式预警][{level}] {title}\n"
                        f"文末参考文献列表中未找到 {pattern} 的对应条目。\n"
                        f"建议：{sugg}"
                    )
                    self._add_comment(doc, para, comment)
                    annotated.add(gn)

        if not annotated and paragraphs:
            # Fallback: comment on first paragraph
            comment = self._fmt_comment(level, title, "引用完整性", "", desc, sugg)
            self._add_comment(doc, paragraphs[0], comment)

    def _handle_punct_error(self, doc, paragraphs: list, error: dict) -> None:
        """Find the sentence via the excerpt in desc, mark red + comment."""
        level = error.get("level", "L2")
        title = error.get("title", "")
        desc  = error.get("desc",  "")
        sugg  = error.get("suggestion", "")

        # desc pattern: "原文：…some text…"
        m = re.search(r"原文[：:]\s*[…\.]{0,3}([^…\.]{4,45})[…\.]?", desc)
        if m:
            needle = m.group(1).strip()
            for para in paragraphs:
                if needle in para.text:
                    self._add_red_text_marker(para)
                    comment = (
                        f"[格式预警][{level}] {title}\n"
                        f"{desc}\n"
                        f"建议：{sugg}"
                    )
                    self._add_comment(doc, para, comment)
                    return

        self._handle_generic(doc, paragraphs, error)

    def _handle_generic(self, doc, paragraphs: list, error: dict) -> None:
        """Insert a layout/structure comment at the best matching paragraph."""
        level = error.get("level", "L3")
        title = error.get("title", "")
        desc  = error.get("desc",  "")
        sugg  = error.get("suggestion", "")
        etype = error.get("type", "")
        loc   = error.get("location", "")

        comment = self._fmt_comment(level, title, etype, loc, desc, sugg)
        target  = self._find_target_paragraph(paragraphs, loc)
        if target:
            self._add_comment(doc, target, comment)

    def _find_target_paragraph(self, paragraphs: list, location_hint) -> object | None:
        """Heuristically pick the best paragraph for a generic comment."""
        if not paragraphs:
            return None

        loc_str = location_hint if isinstance(location_hint, str) else str(location_hint)

        kw_map = {
            "摘要":     ["摘要", "Abstract"],
            "参考文献": ["参考文献", "References"],
            "关键词":   ["关键词", "Keywords"],
            "目录":     ["目录", "Contents", "Table of"],
        }
        for kw, searches in kw_map.items():
            if kw in loc_str:
                for p in paragraphs:
                    if any(s in p.text for s in searches):
                        return p

        # First non-empty paragraph
        for p in paragraphs[:8]:
            if p.text.strip():
                return p

        return paragraphs[0]

    # ------------------------------------------------------------------
    # Low-level XML helpers
    # ------------------------------------------------------------------

    def _add_red_text_marker(
        self,
        paragraph,
        start_idx: int | None = None,
        end_idx:   int | None = None,
    ) -> None:
        """Make runs red + bold (optionally restricted to a char range)."""
        char_pos = 0
        for run in paragraph.runs:
            run_len  = len(run.text)
            in_range = True
            if start_idx is not None and end_idx is not None:
                in_range = char_pos < end_idx and (char_pos + run_len) > start_idx
            if in_range:
                r_el = run._r
                rpr  = r_el.find(_w("rPr"))
                if rpr is None:
                    rpr = etree.Element(_w("rPr"))
                    r_el.insert(0, rpr)

                color_el = rpr.find(_w("color"))
                if color_el is None:
                    color_el = etree.SubElement(rpr, _w("color"))
                color_el.set(_w("val"), "FF0000")

                if rpr.find(_w("b")) is None:
                    etree.SubElement(rpr, _w("b"))

            char_pos += run_len

    def _add_yellow_highlight_to_para(
        self, paragraph, search_text: str | None = None
    ) -> None:
        """Add yellow highlight to runs (optionally only those containing search_text)."""
        for run in paragraph.runs:
            if search_text and search_text not in run.text:
                continue
            r_el = run._r
            rpr  = r_el.find(_w("rPr"))
            if rpr is None:
                rpr = etree.Element(_w("rPr"))
                r_el.insert(0, rpr)
            hl = rpr.find(_w("highlight"))
            if hl is None:
                hl = etree.SubElement(rpr, _w("highlight"))
            hl.set(_w("val"), "yellow")

    def _add_comment(
        self, doc, paragraph, comment_text: str, author: str = "格式校验"
    ) -> None:
        """
        Attach a Word comment to a paragraph by inserting
        w:commentRangeStart / w:commentRangeEnd / w:commentReference
        into the paragraph's lxml element.
        """
        cid = self._alloc_id()
        self._comment_records.append(
            {
                "id":     cid,
                "author": author,
                "date":   datetime.utcnow().strftime("%Y-%m-%dT%H:%M:%SZ"),
                "text":   comment_text,
            }
        )

        p_el = paragraph._p
        runs  = p_el.findall(_w("r"))

        crs = etree.Element(_w("commentRangeStart"))
        crs.set(_w("id"), str(cid))

        cre = etree.Element(_w("commentRangeEnd"))
        cre.set(_w("id"), str(cid))

        ref_run = etree.Element(_w("r"))
        ref_rpr = etree.SubElement(ref_run, _w("rPr"))
        rs_el   = etree.SubElement(ref_rpr, _w("rStyle"))
        rs_el.set(_w("val"), "CommentReference")
        cr_el = etree.SubElement(ref_run, _w("commentReference"))
        cr_el.set(_w("id"), str(cid))

        if runs:
            children      = list(p_el)
            first_run_idx = children.index(runs[0])
            last_run_idx  = children.index(runs[-1])
            p_el.insert(first_run_idx, crs)
            # Indices shifted +1 after the insert above
            p_el.insert(last_run_idx + 2, cre)
            p_el.insert(last_run_idx + 3, ref_run)
        else:
            p_el.append(crs)
            p_el.append(cre)
            p_el.append(ref_run)

    def _alloc_id(self) -> int:
        self._next_id += 1
        return self._next_id

    @staticmethod
    def _fmt_comment(
        level: str, title: str, etype: str, location, desc: str, sugg: str
    ) -> str:
        loc_str = location if isinstance(location, str) else str(location)
        return (
            f"[格式预警][{level}] {title}\n"
            f"类型：{etype}　位置：{loc_str}\n"
            f"实际：{desc}\n"
            f"建议：{sugg}"
        )

    # ------------------------------------------------------------------
    # ZIP post-processing
    # ------------------------------------------------------------------

    def _inject_comments_into_zip(self, docx_bytes: bytes) -> bytes:
        """Inject comments.xml into the saved docx ZIP and patch manifests."""
        comments_xml = self._build_comments_xml()

        in_buf  = io.BytesIO(docx_bytes)
        out_buf = io.BytesIO()

        with zipfile.ZipFile(in_buf, "r") as zin, zipfile.ZipFile(
            out_buf, "w", zipfile.ZIP_DEFLATED
        ) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)

                if item.filename == "[Content_Types].xml":
                    data = self._patch_content_types(data)
                elif item.filename in (
                    "word/_rels/document.xml.rels",
                    "word/_rels/document.xml.rels/",
                ):
                    data = self._patch_rels(data)

                zout.writestr(item, data)

            # Add the comments part
            zout.writestr("word/comments.xml", comments_xml)

        out_buf.seek(0)
        return out_buf.read()

    def _build_comments_xml(self) -> bytes:
        """Build a valid word/comments.xml from collected comment records."""
        root = etree.Element(_w("comments"), nsmap=NSMAP)

        for c in self._comment_records:
            comment_el = etree.SubElement(root, _w("comment"))
            comment_el.set(_w("id"),       str(c["id"]))
            comment_el.set(_w("author"),   c["author"])
            comment_el.set(_w("date"),     c["date"])
            comment_el.set(_w("initials"), "FK")

            p_el = etree.SubElement(comment_el, _w("p"))

            ppr   = etree.SubElement(p_el, _w("pPr"))
            pstyle = etree.SubElement(ppr, _w("pStyle"))
            pstyle.set(_w("val"), "CommentText")

            # Annotation-reference run (required by Word)
            r0  = etree.SubElement(p_el, _w("r"))
            rpr = etree.SubElement(r0, _w("rPr"))
            rs  = etree.SubElement(rpr, _w("rStyle"))
            rs.set(_w("val"), "CommentReference")
            etree.SubElement(r0, _w("annotationRef"))

            # Actual text (line-breaks become <w:br/>)
            lines = c["text"].split("\n")
            for i, line in enumerate(lines):
                r = etree.SubElement(p_el, _w("r"))
                if i > 0:
                    etree.SubElement(r, _w("br"))
                t = etree.SubElement(r, _w("t"))
                t.text = line
                if line != line.strip():
                    t.set(XML_SPACE, "preserve")

        header = b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        return header + etree.tostring(root, encoding="unicode").encode("utf-8")

    @staticmethod
    def _patch_content_types(data: bytes) -> bytes:
        """Insert the comments Override entry if absent."""
        override = (
            b'<Override PartName="/word/comments.xml" ContentType="'
            + COMMENTS_CT.encode()
            + b'"/>'
        )
        if b"comments.xml" in data:
            return data
        return data.replace(b"</Types>", override + b"</Types>")

    @staticmethod
    def _patch_rels(data: bytes) -> bytes:
        """Insert a Relationship for comments.xml if absent."""
        rel = (
            b'<Relationship Id="rId_fmt_comments" '
            b'Type="http://schemas.openxmlformats.org/officeDocument/'
            b'2006/relationships/comments" '
            b'Target="comments.xml"/>'
        )
        if b"comments.xml" in data:
            return data
        return data.replace(b"</Relationships>", rel + b"</Relationships>")
