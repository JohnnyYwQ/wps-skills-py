# -*- coding: utf-8 -*-
"""
linux_word.py — Linux 平台 Word 后端（纯标准库 zipfile + xml.etree 实现 .docx 读写）

架构：与 Windows COM 后端同一套 action 契约，底层直接读写 OpenXML 文件。
- "活动文档" = 内存中的 docx 包（dict: part_name -> bytes）
- createDocument 构建最小合法 docx（styles/document）
- insertText/insertTable/insertImage 等操作修改内存包，save 落盘

单位换算：COM 用磅(pt)，OOXML 用半磅(sz)/缇(twips, 1pt=20twips)
依赖：仅 Python 标准库。x86 / ARM 通用。
"""
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from linux_common import WpsCli

# ---------- OpenXML 命名空间 ----------
NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"
NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

for prefix, uri in (("w", NS_W), ("r", NS_R), ("rel", NS_REL),
                    ("wp", NS_WP), ("a", NS_A), ("pic", NS_PIC)):
    ET.register_namespace(prefix, uri)


def _q(ns, tag):
    return f"{{{ns}}}{tag}"


def _esc(s) -> str:
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


_ALIGN = {"left": "left", "center": "center", "right": "right", "justify": "both"}


# ---------- 最小合法 docx 模板 ----------

def _styles_xml() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:styles xmlns:w="' + NS_W + '">'
            '<w:docDefaults><w:rPrDefault><w:rPr>'
            '<w:rFonts w:ascii="Calibri" w:eastAsia="宋体" w:hAnsi="Calibri" w:cs="Times New Roman"/>'
            '<w:sz w:val="22"/><w:szCs w:val="22"/>'
            '</w:rPr></w:rPrDefault><w:pPrDefault/></w:docDefaults>'
            '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            '<w:name w:val="Normal"/><w:qFormat/></w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading1">'
            '<w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:qFormat/>'
            '<w:pPr><w:keepNext/><w:outlineLvl w:val="0"/></w:pPr>'
            '<w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading2">'
            '<w:name w:val="heading 2"/><w:basedOn w:val="Normal"/><w:qFormat/>'
            '<w:pPr><w:keepNext/><w:outlineLvl w:val="1"/></w:pPr>'
            '<w:rPr><w:b/><w:sz w:val="26"/></w:rPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="Heading3">'
            '<w:name w:val="heading 3"/><w:basedOn w:val="Normal"/><w:qFormat/>'
            '<w:pPr><w:keepNext/><w:outlineLvl w:val="2"/></w:pPr>'
            '<w:rPr><w:b/><w:sz w:val="24"/></w:rPr></w:style>'
            '<w:style w:type="paragraph" w:styleId="Title">'
            '<w:name w:val="Title"/><w:basedOn w:val="Normal"/><w:qFormat/>'
            '<w:rPr><w:b/><w:sz w:val="56"/></w:rPr></w:style>'
            '</w:styles>')


def _document_xml() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:document xmlns:w="' + NS_W + '" xmlns:r="' + NS_R + '">'
            '<w:body>'
            '<w:sectPr>'
            '<w:pgSz w:w="11906" w:h="16838"/>'
            '<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800" '
            'w:header="851" w:footer="992" w:gutter="0"/>'
            '</w:sectPr>'
            '</w:body>'
            '</w:document>')


def _content_types_xml(has_styles=True, images=(), headers=(), footers=()) -> str:
    overrides = [
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>',
    ]
    if has_styles:
        overrides.append('<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>')
    for n in headers:
        overrides.append('<Override PartName="/word/header{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'.format(n=n))
    for n in footers:
        overrides.append('<Override PartName="/word/footer{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'.format(n=n))
    defaults = ('<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                '<Default Extension="xml" ContentType="application/xml"/>')
    exts = set()
    for _, ext in images:
        exts.add(ext.lower())
    mime = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg",
            "gif": "image/gif", "bmp": "image/bmp", "wmf": "image/x-wmf"}
    for e in exts:
        defaults += '<Default Extension="{}" ContentType="{}"/>'.format(e, mime.get(e, "application/octet-stream"))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="' + NS_CT + '">' + defaults + "".join(overrides) + '</Types>')


def _root_rels_xml() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="' + NS_REL + '">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
            '</Relationships>')


def _doc_rels_xml(images=(), styles=True) -> str:
    rels = []
    rid = 1
    if styles:
        rels.append('<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>')
        rid = 2
    for i, (fname, _) in enumerate(images):
        rels.append('<Relationship Id="rId{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{fname}"/>'.format(rid=rid + i, fname=fname))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="' + NS_REL + '">' + "".join(rels) + '</Relationships>')


# ---------- docx 包操作核心 ----------

class LinuxWordBridge:
    """纯标准库 docx 读写后端"""

    def __init__(self):
        self._parts = {}
        self._path = None
        self._next_rid = 10   # document.xml.rels 的图片 rId 计数
        self._next_img = 1
        self._cli = WpsCli()

    # ----- 包构建 -----

    def _ensure_doc(self):
        if "word/document.xml" not in self._parts:
            raise RuntimeError("没有活动文档，请先 createDocument 或 openDocument")

    def _body(self):
        """解析并返回 document.xml 的 body Element（每次现解析，改完写回）"""
        root = ET.fromstring(self._parts["word/document.xml"])
        body = root.find(_q(NS_W, "body"))
        if body is None:
            raise RuntimeError("文档结构异常：无 body")
        return body

    def _write_body(self, body):
        # body 的父节点 document 需要保留，重新序列化整个树
        root = ET.fromstring(self._parts["word/document.xml"])
        old = root.find(_q(NS_W, "body"))
        root.remove(old)
        root.append(body)
        self._parts["word/document.xml"] = ET.tostring(root, encoding="unicode",
                                                       xml_declaration=True)

    @staticmethod
    def _new_paragraph(text="", style=None, align=None, font_name=None,
                       font_size=None, bold=None, italic=None, color=None,
                       underline=None, line_spacing=None):
        """构建一个 w:p Element"""
        p = ET.Element(_q(NS_W, "p"))
        pPr = None
        if style or align or line_spacing:
            pPr = ET.SubElement(p, _q(NS_W, "pPr"))
            if style:
                ps = ET.SubElement(pPr, _q(NS_W, "pStyle"))
                ps.set(_q(NS_W, "val"), style)
            if line_spacing:
                sp = ET.SubElement(pPr, _q(NS_W, "spacing"))
                sp.set(_q(NS_W, "line"), str(int(float(line_spacing) * 20)))
                sp.set(_q(NS_W, "lineRule"), "exact")
            if align:
                jc = ET.SubElement(pPr, _q(NS_W, "jc"))
                jc.set(_q(NS_W, "val"), _ALIGN.get(align, align))
        lines = str(text).split("\n") if text else [""]
        for i, line in enumerate(lines):
            if i > 0:
                br_p = ET.SubElement(p, _q(NS_W, "r"))
                ET.SubElement(br_p, _q(NS_W, "br"))
            r = ET.SubElement(p, _q(NS_W, "r"))
            rPr = None
            if font_name:
                rPr = ET.SubElement(r, _q(NS_W, "rPr"))
                fonts = ET.SubElement(rPr, _q(NS_W, "rFonts"))
                fonts.set(_q(NS_W, "ascii"), font_name)
                fonts.set(_q(NS_W, "eastAsia"), font_name)
                fonts.set(_q(NS_W, "hAnsi"), font_name)
            if font_size:
                if rPr is None:
                    rPr = ET.SubElement(r, _q(NS_W, "rPr"))
                sz = ET.SubElement(rPr, _q(NS_W, "sz"))
                sz.set(_q(NS_W, "val"), str(int(float(font_size) * 2)))
            if bold is not None:
                if rPr is None:
                    rPr = ET.SubElement(r, _q(NS_W, "rPr"))
                ET.SubElement(rPr, _q(NS_W, "b"))
            if italic is not None:
                if rPr is None:
                    rPr = ET.SubElement(r, _q(NS_W, "rPr"))
                ET.SubElement(rPr, _q(NS_W, "i"))
            if underline is not None:
                if rPr is None:
                    rPr = ET.SubElement(r, _q(NS_W, "rPr"))
                ET.SubElement(rPr, _q(NS_W, "u"))
            if color:
                if rPr is None:
                    rPr = ET.SubElement(r, _q(NS_W, "rPr"))
                c = ET.SubElement(rPr, _q(NS_W, "color"))
                c.set(_q(NS_W, "val"), str(color).replace("#", "").upper())
            t = ET.SubElement(r, _q(NS_W, "t"))
            t.text = line
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        return p

    def _append_to_body(self, elem):
        body = self._body()
        # sectPr 必须保持在最后
        sect = body.find(_q(NS_W, "sectPr"))
        if sect is not None:
            body.remove(sect)
            body.append(elem)
            body.append(sect)
        else:
            body.append(elem)
        self._write_body(body)

    # ----- action 实现 -----

    def ping(self, params):
        return {"success": True, "data": {"message": "pong", "backend": "stdlib-openxml"}}

    def getAppInfo(self, params):
        return {"success": True, "data": {
            "app": "WPS文字(Linux文件模式)",
            "backend": "stdlib OpenXML",
            "wpsCli": self._cli.info(),
        }}

    def createDocument(self, params):
        self._parts = {
            "[Content_Types].xml": _content_types_xml(),
            "_rels/.rels": _root_rels_xml(),
            "word/document.xml": _document_xml(),
            "word/_rels/document.xml.rels": _doc_rels_xml(),
            "word/styles.xml": _styles_xml(),
        }
        self._path = params.get("filePath")
        return {"success": True, "data": {"name": "文档1", "paragraphs": 0,
                                         "path": self._path or "(未保存)"}}

    def openDocument(self, params):
        fp = params.get("filePath")
        if not fp:
            return {"success": False, "error": "缺少 filePath"}
        fp = os.path.abspath(fp)
        if not os.path.isfile(fp):
            return {"success": False, "error": f"文件不存在: {fp}"}
        try:
            self._parts = {}
            with zipfile.ZipFile(fp, "r") as z:
                for name in z.namelist():
                    self._parts[name] = z.read(name)
            self._path = fp
        except Exception as e:
            return {"success": False, "error": f"打开失败（仅支持 .docx）: {e}"}
        if params.get("openInWps"):
            self._cli.open_file("word", fp)
        body = self._body()
        paras = len(body.findall(_q(NS_W, "p")))
        return {"success": True, "data": {"name": os.path.basename(fp),
                                         "paragraphs": paras}}

    def save(self, params):
        self._ensure_doc()
        if not self._path:
            return {"success": False, "error": "文档未指定保存路径，请用 saveAs"}
        self._write_zip(self._path)
        return {"success": True, "data": {"path": self._path}}

    def saveAs(self, params):
        self._ensure_doc()
        fp = params.get("filePath")
        if not fp:
            return {"success": False, "error": "缺少 filePath"}
        fp = os.path.abspath(fp)
        os.makedirs(os.path.dirname(fp), exist_ok=True)
        self._write_zip(fp)
        self._path = fp
        return {"success": True, "data": {"path": fp}}

    def _write_zip(self, fp):
        # 重建 Content_Types / rels（图片可能有增删）
        images = []
        for name in self._parts:
            m = re.match(r"^word/media/(img\d+)\.(\w+)$", name)
            if m:
                images.append((m.group(1) + "." + m.group(2), m.group(2)))
        self._parts["[Content_Types].xml"] = _content_types_xml(
            has_styles="word/styles.xml" in self._parts, images=images)
        # document.xml.rels：styles + images
        rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>']
        for i, (fname, _) in enumerate(images):
            rels.append('<Relationship Id="rId{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{fname}"/>'.format(rid=i + 2, fname=fname))
        self._parts["word/_rels/document.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="' + NS_REL + '">' + "".join(rels) + '</Relationships>')
        with zipfile.ZipFile(fp, "w", zipfile.ZIP_DEFLATED) as z:
            z.writestr("[Content_Types].xml", self._parts["[Content_Types].xml"])
            for name, content in self._parts.items():
                if name == "[Content_Types].xml":
                    continue
                if isinstance(content, str):
                    content = content.encode("utf-8")
                z.writestr(name, content)

    def getActiveDocument(self, params):
        self._ensure_doc()
        body = self._body()
        paras = len(body.findall(_q(NS_W, "p")))
        return {"success": True, "data": {
            "name": os.path.basename(self._path) if self._path else "文档1",
            "paragraphs": paras,
        }}

    def getOpenDocuments(self, params):
        if self._path:
            return {"success": True, "data": {"documents": [os.path.basename(self._path)],
                                             "count": 1}}
        return {"success": True, "data": {"documents": [], "count": 0}}

    def insertText(self, params):
        self._ensure_doc()
        text = params.get("text", "")
        style = params.get("style")
        # style 名映射：中文/常见名 -> styleId
        style_map = {"标题1": "Heading1", "标题 1": "Heading1", "Heading1": "Heading1",
                     "标题2": "Heading2", "标题 2": "Heading2", "Heading2": "Heading2",
                     "标题3": "Heading3", "标题 3": "Heading3", "Heading3": "Heading3",
                     "标题": "Title", "Title": "Title", "正文": "Normal", "Normal": "Normal"}
        style_id = style_map.get(style, style)
        align = params.get("alignment")
        p = self._new_paragraph(
            text=text, style=style_id, align=align,
            font_name=params.get("font_name"),
            font_size=params.get("font_size"),
            bold=params.get("bold") if params.get("bold") is not None else None,
        )
        self._append_to_body(p)
        return {"success": True, "data": {"inserted": True}}

    def getDocumentText(self, params):
        self._ensure_doc()
        body = self._body()
        chunks = []
        for elem in body:
            if elem.tag == _q(NS_W, "p"):
                text = "".join(t.text or "" for t in elem.iter(_q(NS_W, "t")))
                chunks.append(text)
            elif elem.tag == _q(NS_W, "tbl"):
                for row in elem.iter(_q(NS_W, "tr")):
                    cells = []
                    for tc in row.findall(_q(NS_W, "tc")):
                        cells.append("".join(t.text or "" for t in tc.iter(_q(NS_W, "t"))))
                    chunks.append("\t".join(cells))
        full = "\n".join(chunks)
        start = params.get("start")
        end = params.get("end")
        if start is not None or end is not None:
            s = int(start) if start else 0
            e = int(end) if end else len(full)
            full = full[s:e]
        return {"success": True, "data": {"text": full, "length": len(full)}}

    def getDocumentParagraphs(self, params):
        self._ensure_doc()
        body = self._body()
        paras = body.findall(_q(NS_W, "p"))
        s = int(params.get("start_paragraph", 1))
        e = int(params.get("end_paragraph", len(paras)))
        out = []
        for i in range(s, min(e, len(paras)) + 1):
            p = paras[i - 1]
            text = "".join(t.text or "" for t in p.iter(_q(NS_W, "t")))
            pPr = p.find(_q(NS_W, "pPr"))
            style = "Normal"
            if pPr is not None:
                ps = pPr.find(_q(NS_W, "pStyle"))
                if ps is not None:
                    style = ps.get(_q(NS_W, "val"), "Normal")
            out.append({"index": i, "text": text, "style": style})
        return {"success": True, "data": {"paragraphs": out, "count": len(out)}}

    def setFont(self, params):
        """range: all / last / N(段落序号)"""
        self._ensure_doc()
        body = self._body()
        paras = body.findall(_q(NS_W, "p"))
        rng = params.get("range", "all")
        if rng == "last":
            targets = [paras[-1]] if paras else []
        elif isinstance(rng, int) or (isinstance(rng, str) and rng.isdigit()):
            idx = int(rng)
            targets = [paras[idx - 1]] if 1 <= idx <= len(paras) else []
        else:
            targets = paras
        for p in targets:
            for r in p.findall(_q(NS_W, "r")):
                rPr = r.find(_q(NS_W, "rPr"))
                if rPr is None:
                    rPr = ET.Element(_q(NS_W, "rPr"))
                    r.insert(0, rPr)
                if params.get("font_name"):
                    fonts = rPr.find(_q(NS_W, "rFonts"))
                    if fonts is None:
                        fonts = ET.SubElement(rPr, _q(NS_W, "rFonts"))
                    for k in ("ascii", "eastAsia", "hAnsi"):
                        fonts.set(_q(NS_W, k), params["font_name"])
                if params.get("font_size"):
                    for tag in ("sz", "szCs"):
                        e = rPr.find(_q(NS_W, tag))
                        if e is None:
                            e = ET.SubElement(rPr, _q(NS_W, tag))
                        e.set(_q(NS_W, "val"), str(int(float(params["font_size"]) * 2)))
                if params.get("bold") is not None:
                    if params["bold"]:
                        if rPr.find(_q(NS_W, "b")) is None:
                            ET.SubElement(rPr, _q(NS_W, "b"))
                    else:
                        for b in rPr.findall(_q(NS_W, "b")):
                            rPr.remove(b)
                if params.get("italic") is not None:
                    if params["italic"]:
                        if rPr.find(_q(NS_W, "i")) is None:
                            ET.SubElement(rPr, _q(NS_W, "i"))
                if params.get("underline") is not None and params["underline"]:
                    if rPr.find(_q(NS_W, "u")) is None:
                        ET.SubElement(rPr, _q(NS_W, "u"))
                if params.get("color"):
                    c = rPr.find(_q(NS_W, "color"))
                    if c is None:
                        c = ET.SubElement(rPr, _q(NS_W, "color"))
                    c.set(_q(NS_W, "val"), str(params["color"]).replace("#", "").upper())
        self._write_body(body)
        return {"success": True, "data": {"target": rng, "runs": len(targets)}}

    def setTextColor(self, params):
        return self.setFont({**params, "color": params.get("color", "FF0000")})

    def setLineSpacing(self, params):
        self._ensure_doc()
        body = self._body()
        paras = body.findall(_q(NS_W, "p"))
        idx = params.get("paragraphIndex")
        targets = [paras[int(idx) - 1]] if idx and 1 <= int(idx) <= len(paras) else paras
        spacing = params.get("spacing", 1.5)
        for p in targets:
            pPr = p.find(_q(NS_W, "pPr"))
            if pPr is None:
                pPr = ET.Element(_q(NS_W, "pPr"))
                p.insert(0, pPr)
            sp = pPr.find(_q(NS_W, "spacing"))
            if sp is None:
                sp = ET.SubElement(pPr, _q(NS_W, "spacing"))
            sp.set(_q(NS_W, "line"), str(int(float(spacing) * 240)))
            sp.set(_q(NS_W, "lineRule"), "auto")
        self._write_body(body)
        return {"success": True}

    def applyStyle(self, params):
        self._ensure_doc()
        body = self._body()
        paras = body.findall(_q(NS_W, "p"))
        idx = params.get("paragraphIndex")
        targets = [paras[int(idx) - 1]] if idx and 1 <= int(idx) <= len(paras) else paras
        style_map = {"标题1": "Heading1", "标题 1": "Heading1", "Heading1": "Heading1",
                     "标题2": "Heading2", "标题 2": "Heading2", "Heading2": "Heading2",
                     "标题3": "Heading3", "标题 3": "Heading3", "Heading3": "Heading3",
                     "标题": "Title", "Title": "Title", "正文": "Normal", "Normal": "Normal"}
        style_id = style_map.get(params.get("style"), params.get("style", "Normal"))
        for p in targets:
            pPr = p.find(_q(NS_W, "pPr"))
            if pPr is None:
                pPr = ET.Element(_q(NS_W, "pPr"))
                p.insert(0, pPr)
            ps = pPr.find(_q(NS_W, "pStyle"))
            if ps is None:
                ps = ET.SubElement(pPr, _q(NS_W, "pStyle"))
            ps.set(_q(NS_W, "val"), style_id)
        self._write_body(body)
        return {"success": True, "data": {"style": style_id}}

    def setParagraph(self, params):
        self._ensure_doc()
        body = self._body()
        paras = body.findall(_q(NS_W, "p"))
        idx = params.get("paragraphIndex")
        p = paras[int(idx) - 1] if idx and 1 <= int(idx) <= len(paras) else (paras[-1] if paras else None)
        if p is None:
            return {"success": False, "error": "无段落"}
        pPr = p.find(_q(NS_W, "pPr"))
        if pPr is None:
            pPr = ET.Element(_q(NS_W, "pPr"))
            p.insert(0, pPr)
        if params.get("alignment"):
            jc = pPr.find(_q(NS_W, "jc"))
            if jc is None:
                jc = ET.SubElement(pPr, _q(NS_W, "jc"))
            jc.set(_q(NS_W, "val"), _ALIGN.get(params["alignment"], params["alignment"]))
        if params.get("lineSpacing"):
            sp = pPr.find(_q(NS_W, "spacing"))
            if sp is None:
                sp = ET.SubElement(pPr, _q(NS_W, "spacing"))
            sp.set(_q(NS_W, "line"), str(int(float(params["lineSpacing"]) * 240)))
            sp.set(_q(NS_W, "lineRule"), "auto")
        if params.get("spaceBefore"):
            sp = pPr.find(_q(NS_W, "spacing"))
            if sp is None:
                sp = ET.SubElement(pPr, _q(NS_W, "spacing"))
            sp.set(_q(NS_W, "before"), str(int(float(params["spaceBefore"]) * 20)))
        if params.get("spaceAfter"):
            sp = pPr.find(_q(NS_W, "spacing"))
            if sp is None:
                sp = ET.SubElement(pPr, _q(NS_W, "spacing"))
            sp.set(_q(NS_W, "after"), str(int(float(params["spaceAfter"]) * 20)))
        self._write_body(body)
        return {"success": True}

    def findReplace(self, params):
        self._ensure_doc()
        find_t = params.get("find_text", "")
        rep_t = params.get("replace_text", "")
        if not find_t:
            return {"success": False, "error": "缺少 find_text"}
        count = 0
        body = self._body()
        # 简化策略：逐 run 内替换（跨 run 的匹配不处理）
        for t in body.iter(_q(NS_W, "t")):
            if t.text and find_t in t.text:
                t.text = t.text.replace(find_t, rep_t)
                count += t.text.count(rep_t) if rep_t else 0
        self._write_body(body)
        return {"success": True, "data": {"replaced": count}}

    def findInDocument(self, params):
        self._ensure_doc()
        find_t = params.get("find_text", "")
        body = self._body()
        results = []
        pos = 0
        for elem in body.iter():
            if elem.tag == _q(NS_W, "t") and elem.text:
                idx = 0
                while True:
                    i = elem.text.find(find_t, idx)
                    if i < 0:
                        break
                    results.append({"start": pos + i, "end": pos + i + len(find_t),
                                    "text": find_t})
                    idx = i + len(find_t)
                pos += len(elem.text)
        return {"success": True, "data": {"results": results, "count": len(results)}}

    def insertTable(self, params):
        self._ensure_doc()
        rows = int(params.get("rows", 3))
        cols = int(params.get("cols", 3))
        data = params.get("data")
        tbl = ET.Element(_q(NS_W, "tbl"))
        # 表属性：全边框
        tblPr = ET.SubElement(tbl, _q(NS_W, "tblPr"))
        borders = ET.SubElement(tblPr, _q(NS_W, "tblBorders"))
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            b = ET.SubElement(borders, _q(NS_W, side))
            b.set(_q(NS_W, "val"), "single")
            b.set(_q(NS_W, "sz"), "4")
            b.set(_q(NS_W, "color"), "auto")
        tblW = ET.SubElement(tblPr, _q(NS_W, "tblW"))
        tblW.set(_q(NS_W, "w"), "5000")
        tblW.set(_q(NS_W, "type"), "pct")
        for r in range(rows):
            tr = ET.SubElement(tbl, _q(NS_W, "tr"))
            for c in range(cols):
                tc = ET.SubElement(tr, _q(NS_W, "tc"))
                text = ""
                if data and r < len(data) and c < len(data[r]):
                    text = str(data[r][c])
                elif r == 0:
                    text = f"表头{c + 1}"
                tcPr = ET.SubElement(tc, _q(NS_W, "tcPr"))
                tcW = ET.SubElement(tcPr, _q(NS_W, "tcW"))
                tcW.set(_q(NS_W, "w"), str(int(5000 / cols)))
                tcW.set(_q(NS_W, "type"), "pct")
                p = ET.SubElement(tc, _q(NS_W, "p"))
                run = ET.SubElement(p, _q(NS_W, "r"))
                if r == 0:
                    rPr = ET.SubElement(run, _q(NS_W, "rPr"))
                    ET.SubElement(rPr, _q(NS_W, "b"))
                t = ET.SubElement(run, _q(NS_W, "t"))
                t.text = text
        self._append_to_body(tbl)
        # 表后补一个空段落（Word 规范要求表格后有段落）
        self._append_to_body(self._new_paragraph())
        return {"success": True, "data": {"rows": rows, "cols": cols}}

    def insertImage(self, params):
        self._ensure_doc()
        img_path = params.get("imagePath") or params.get("filePath")
        if not img_path or not os.path.isfile(img_path):
            return {"success": False, "error": f"图片不存在: {img_path}"}
        ext = os.path.splitext(img_path)[1].lstrip(".").lower()
        if ext not in ("png", "jpg", "jpeg", "gif", "bmp"):
            return {"success": False, "error": f"不支持的图片格式: {ext}"}
        # 读取图片并获取尺寸（纯 Python 解析 PNG/JPEG 头，不需要 Pillow）
        from linux_common import setup_vendor_path  # noqa
        w, h = _read_image_size(img_path)
        # 归一到合理显示宽度（EMU: 1cm=360000, 最大约 15cm）
        disp_w = params.get("width")
        disp_h = params.get("height")
        if disp_w and disp_h:
            emu_w, emu_h = int(disp_w) * 9525, int(disp_h) * 9525  # px->EMU
        elif w and h:
            scale = min(1.0, 5400000 / w)
            emu_w, emu_h = int(w * scale), int(h * scale)
        else:
            emu_w, emu_h = 3600000, 2400000
        # 存入包
        fname = f"img{self._next_img}.{ext}"
        self._next_img += 1
        with open(img_path, "rb") as f:
            self._parts[f"word/media/{fname}"] = f.read()
        rid = f"rId{self._next_rid}"
        self._next_rid += 1
        # 构建 drawing XML
        drawing = (
            '<w:r xmlns:w="' + NS_W + '" xmlns:r="' + NS_R + '">'
            '<w:drawing>'
            '<wp:inline xmlns:wp="' + NS_WP + '" distT="0" distB="0" distL="0" distR="0">'
            '<wp:extent cx="{w}" cy="{h}"/>'
            '<wp:docPr id="{id}" name="Picture {id}"/>'
            '<a:graphic xmlns:a="' + NS_A + '">'
            '<a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
            '<pic:pic xmlns:pic="' + NS_PIC + '">'
            '<pic:nvPicPr>'
            '<pic:cNvPr id="{id}" name="Picture {id}"/>'
            '<pic:cNvPicPr/>'
            '</pic:nvPicPr>'
            '<pic:blipFill>'
            '<a:blip r:embed="{rid}"/>'
            '<a:stretch><a:fillRect/></a:stretch>'
            '</pic:blipFill>'
            '<pic:spPr>'
            '<a:xfrm><a:off x="0" y="0"/><a:ext cx="{w}" cy="{h}"/></a:xfrm>'
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
            '</pic:spPr>'
            '</pic:pic>'
            '</a:graphicData>'
            '</a:graphic>'
            '</wp:inline>'
            '</w:drawing>'
            '</w:r>').format(w=emu_w, h=emu_h, id=self._next_img, rid=rid)
        # 把 run 追加到最后一个段落，或新建段落
        body = self._body()
        paras = body.findall(_q(NS_W, "p"))
        run = ET.fromstring(drawing)
        if paras and body[-1].tag != _q(NS_W, "sectPr"):
            paras[-1].append(run)
        else:
            p = ET.Element(_q(NS_W, "p"))
            p.append(run)
            self._append_to_body(p)
        # document.xml.rels 增加图片关系
        rels_name = "word/_rels/document.xml.rels"
        content = self._parts.get(rels_name, _doc_rels_xml())
        if isinstance(content, bytes):
            content = content.decode("utf-8")
        content = content.replace(
            "</Relationships>",
            '<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/{fname}"/></Relationships>'.format(rid=rid, fname=fname))
        self._parts[rels_name] = content
        self._write_body(body)
        return {"success": True, "data": {"width": emu_w, "height": emu_h}}

    def insertPageBreak(self, params):
        self._ensure_doc()
        p = ET.Element(_q(NS_W, "p"))
        r = ET.SubElement(p, _q(NS_W, "r"))
        ET.SubElement(r, _q(NS_W, "br")).set(_q(NS_W, "type"), "page")
        self._append_to_body(p)
        return {"success": True}

    def insertSectionBreak(self, params):
        return self.insertPageBreak(params)

    def insertBookmark(self, params):
        self._ensure_doc()
        name = params.get("name", "bookmark1")
        body = self._body()
        paras = body.findall(_q(NS_W, "p"))
        if not paras:
            return {"success": False, "error": "无段落"}
        p = paras[-1]
        bid = str(len(list(body.iter(_q(NS_W, "bookmarkStart")))) + 1)
        s = ET.Element(_q(NS_W, "bookmarkStart"))
        s.set(_q(NS_W, "id"), bid)
        s.set(_q(NS_W, "name"), name)
        e = ET.Element(_q(NS_W, "bookmarkEnd"))
        e.set(_q(NS_W, "id"), bid)
        p.insert(0, s)
        p.append(e)
        self._write_body(body)
        return {"success": True, "data": {"name": name}}

    def replaceBookmarkContent(self, params):
        self._ensure_doc()
        name = params.get("name")
        body = self._body()
        for s in body.iter(_q(NS_W, "bookmarkStart")):
            if s.get(_q(NS_W, "name")) == name:
                # 找到同段落的 runs，替换文本
                p = None
                for candidate in body.iter(_q(NS_W, "p")):
                    if s in list(candidate):
                        p = candidate
                        break
                if p is None:
                    return {"success": False, "error": "书签所在段落未找到"}
                for r in p.findall(_q(NS_W, "r")):
                    for t in r.iter(_q(NS_W, "t")):
                        t.text = params.get("text", "")
                self._write_body(body)
                return {"success": True}
        return {"success": False, "error": f"书签不存在: {name}"}

    def setPageSetup(self, params):
        self._ensure_doc()
        body = self._body()
        sect = body.find(_q(NS_W, "sectPr"))
        if sect is None:
            sect = ET.SubElement(body, _q(NS_W, "sectPr"))
        if params.get("orientation"):
            sz = sect.find(_q(NS_W, "pgSz"))
            if sz is None:
                sz = ET.SubElement(sect, _q(NS_W, "pgSz"))
            if params["orientation"] == "landscape":
                sz.set(_q(NS_W, "w"), "16838")
                sz.set(_q(NS_W, "h"), "11906")
            else:
                sz.set(_q(NS_W, "w"), "11906")
                sz.set(_q(NS_W, "h"), "16838")
        mar = sect.find(_q(NS_W, "pgMar"))
        if mar is None:
            mar = ET.SubElement(sect, _q(NS_W, "pgMar"))
        for key, attr in (("marginTop", "top"), ("marginBottom", "bottom"),
                          ("marginLeft", "left"), ("marginRight", "right")):
            if params.get(key):
                mar.set(_q(NS_W, attr), str(int(float(params[key]) * 20)))
        self._write_body(body)
        return {"success": True}

    def generateTOC(self, params):
        """插入 TOC 域代码（打开文档后需手动刷新域以生成目录）"""
        self._ensure_doc()
        p = ET.Element(_q(NS_W, "p"))
        fld = ET.SubElement(p, _q(NS_W, "fldSimple"))
        fld.set(_q(NS_W, "instr"), r' TOC \o "1-3" \h \z \u ')
        r = ET.SubElement(fld, _q(NS_W, "r"))
        t = ET.SubElement(r, _q(NS_W, "t"))
        t.text = "（目录域已插入，在 WPS 中按 F9 或右键刷新生成目录）"
        self._append_to_body(p)
        return {"success": True, "data": {"note": "TOC 域已插入，打开后刷新域生成目录"}}

    def addComment(self, params):
        return {"success": False,
                "error": "Linux 文件模式暂不支持批注（需 word/comments.xml 完整链路）"}

    def insertHeader(self, params):
        return self._insert_hf("header", params)

    def insertFooter(self, params):
        return self._insert_hf("footer", params)

    def _insert_hf(self, kind, params):
        """页眉/页脚（简易实现：单 section）"""
        self._ensure_doc()
        text = params.get("text", "")
        num = 1
        part = f"word/{kind}{num}.xml"
        self._parts[part] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:hdr xmlns:w="' + NS_W + '" ' +
            ('type="default"' if kind == "header" else '') + '>'
            if kind == "header" else
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<w:ftr xmlns:w="' + NS_W + '">').replace('>', '>', 1)
        # 重新构建完整 XML（上面 replace 太 hack，直接构建）
        if kind == "header":
            xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<w:hdr xmlns:w="' + NS_W + '">'
                   '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
                   '<w:r><w:t>' + _esc(text) + '</w:t></w:r></w:p>'
                   '</w:hdr>')
        else:
            xml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<w:ftr xmlns:w="' + NS_W + '">'
                   '<w:p><w:pPr><w:jc w:val="center"/></w:pPr>'
                   '<w:r><w:t>' + _esc(text) + '</w:t></w:r></w:p>'
                   '</w:ftr>')
        self._parts[part] = xml
        # document.xml.rels 增加关系
        rid = f"rId{self._next_rid}"
        self._next_rid += 1
        rels_name = "word/_rels/document.xml.rels"
        content = self._parts.get(rels_name, _doc_rels_xml())
        if isinstance(content, bytes):
            content = content.decode("utf-8")
        rel_type = "header" if kind == "header" else "footer"
        content = content.replace(
            "</Relationships>",
            '<Relationship Id="{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/{type}" Target="{kind}1.xml"/></Relationships>'.format(rid=rid, type=rel_type, kind=kind))
        self._parts[rels_name] = content
        # sectPr 增加 headerReference / footerReference
        body = self._body()
        sect = body.find(_q(NS_W, "sectPr"))
        if sect is None:
            sect = ET.SubElement(body, _q(NS_W, "sectPr"))
        ref = ET.Element(_q(NS_W, rel_type + "Reference"))
        ref.set(_q(NS_W, "type"), "default")
        ref.set(_q(NS_R, "id"), rid)
        sect.insert(0, ref)
        self._write_body(body)
        # Content_Types 增加 override
        ct = self._parts["[Content_Types].xml"]
        if isinstance(ct, bytes):
            ct = ct.decode("utf-8")
        if kind + "1.xml" not in ct:
            ct = ct.replace(
                "</Types>",
                '<Override PartName="/word/{kind}1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.{kind}+xml"/></Types>'.format(kind=kind))
            self._parts["[Content_Types].xml"] = ct
        return {"success": True, "data": {"type": kind, "text": text}}

    def convertToPDF(self, params):
        fp = params.get("filePath") or self._path
        if not fp or not os.path.isfile(fp):
            return {"success": False, "error": "缺少 filePath 或文件不存在"}
        if self._path == fp and "word/document.xml" in self._parts:
            self.save({})  # 先保存最新内容
        return self._cli.convert_to_pdf("word", fp, params.get("outputPath"))

    def smartFillField(self, params):
        return {"success": False, "error": "Linux 文件模式暂不支持 smartFillField"}

    # ---------- 统一分发 ----------

    def execute(self, action: str, params: dict) -> dict:
        params = params or {}
        if action == "ping":
            return self.ping(params)
        fn = getattr(self, action, None)
        if fn is None or not callable(fn):
            return {"success": False,
                    "error": f"Linux 文件模式未实现 action '{action}'"}
        try:
            return fn(params)
        except RuntimeError as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": f"{type(e).__name__}: {e}"}


# ---------- 图片尺寸解析（纯 Python，无 Pillow） ----------

def _read_image_size(path: str):
    """读取 PNG/JPEG/GIF/BMP 宽高，失败返回 (None, None)"""
    try:
        with open(path, "rb") as f:
            head = f.read(64)
        if head[:8] == b"\x89PNG\r\n\x1a\n":
            import struct
            w, h = struct.unpack(">II", head[16:24])
            return w, h
        if head[:2] == b"\xff\xd8":  # JPEG
            import struct
            with open(path, "rb") as f:
                f.seek(0)
                while True:
                    marker = f.read(2)
                    if len(marker) < 2 or marker[0] != 0xFF:
                        return None, None
                    if marker[1] in (0xC0, 0xC1, 0xC2, 0xC3):
                        f.read(3)
                        h, w = struct.unpack(">HH", f.read(4))
                        return w, h
                    elif marker[1] in (0xD8, 0xD9):
                        continue
                    else:
                        seg_len = struct.unpack(">H", f.read(2))[0]
                        f.seek(seg_len - 2, 1)
        if head[:6] in (b"GIF87a", b"GIF89a"):
            import struct
            w, h = struct.unpack("<HH", head[6:10])
            return w, h
        if head[:2] == b"BM":
            import struct
            w, h = struct.unpack("<ii", head[18:26])
            return abs(w), abs(h)
    except Exception:
        pass
    return None, None
