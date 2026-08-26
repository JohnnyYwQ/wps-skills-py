# -*- coding: utf-8 -*-
"""
linux_ppt.py — Linux 平台 PPT 后端（纯标准库 zipfile + xml.etree 实现 .pptx 读写）

架构：与 Windows COM 后端同一套 action 契约，底层直接读写 OpenXML 文件。
- "活动演示文稿" = 内存中的 pptx 包（dict: part_name -> bytes）
- createPresentation 构建最小合法 pptx（theme/master/layout/slides）
- addSlide/setSlideTitle/setSlideContent 等操作修改内存包，save 落盘

依赖：仅 Python 标准库。x86 / ARM 通用。
"""
import os
import io
import re
import copy
import zipfile
import xml.etree.ElementTree as ET
from linux_common import WpsCli

# ---------- OpenXML 命名空间 ----------
NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

for prefix, uri in (("p", NS_P), ("a", NS_A), ("r", NS_R), ("rel", NS_REL), ("ct", NS_CT)):
    ET.register_namespace(prefix, uri)


def _q(ns, tag):
    return f"{{{ns}}}{tag}"


# ---------- 最小合法 pptx 模板 ----------

def _theme_xml() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<a:theme xmlns:a="' + NS_A + '" name="Office">'
            '<a:themeElements>'
            '<a:clrScheme name="Office">'
            '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>'
            '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>'
            '<a:dk2><a:srgbClr val="44546A"/></a:dk2>'
            '<a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>'
            '<a:accent1><a:srgbClr val="4472C4"/></a:accent1>'
            '<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>'
            '<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>'
            '<a:accent4><a:srgbClr val="FFC000"/></a:accent4>'
            '<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>'
            '<a:accent6><a:srgbClr val="70AD47"/></a:accent6>'
            '<a:hlink><a:srgbClr val="0563C1"/></a:hlink>'
            '<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>'
            '</a:clrScheme>'
            '<a:fontScheme name="Office">'
            '<a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>'
            '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>'
            '</a:fontScheme>'
            '<a:fmtScheme name="Office">'
            '<a:fillStyleLst>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '</a:fillStyleLst>'
            '<a:lnStyleLst>'
            '<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '<a:ln><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>'
            '</a:lnStyleLst>'
            '<a:effectStyleLst>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '<a:effectStyle><a:effectLst/></a:effectStyle>'
            '</a:effectStyleLst>'
            '<a:bgFillStyleLst>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>'
            '</a:bgFillStyleLst>'
            '</a:fmtScheme>'
            '</a:themeElements>'
            '</a:theme>')


def _slide_master_xml() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sldMaster xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '">'
            '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>'
            '<p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/>'
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title Placeholder"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
            '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Body Placeholder"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="838200" y="1825625"/><a:ext cx="10515600" cy="4351338"/></a:xfrm></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
            '</p:spTree></p:cSld>'
            '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" '
            'accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
            '<p:sldLayoutIdLst>'
            '<p:sldLayoutId id="2147483649" r:id="rId1"/>'
            '</p:sldLayoutIdLst>'
            '</p:sldMaster>')


def _slide_layout_xml() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sldLayout xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '" '
            'type="obj" preserve="1">'
            '<p:cSld name="Title and Content">'
            '<p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/>'
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
            '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
            '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
            '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
            '</p:spTree></p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
            '</p:sldLayout>')


def _new_slide_xml(title: str = "", content: str = "", layout: str = "title_content") -> str:
    """构建单张幻灯片 XML。layout: title / title_content / blank"""
    has_title = layout in ("title", "title_content")
    has_content = layout == "title_content"
    shapes = []
    next_id = 2
    if has_title:
        shapes.append(
            '<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="Title"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="838200" y="365125"/><a:ext cx="10515600" cy="1325563"/></a:xfrm></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="zh-CN"/>'
            '<a:t>{t}</a:t></a:r></a:p></p:txBody></p:sp>'.format(id=next_id, t=_esc(title)))
        next_id += 1
    if has_content:
        paras = "".join(
            '<a:p><a:r><a:rPr lang="zh-CN"/><a:t>{}</a:t></a:r></a:p>'.format(_esc(line))
            for line in (content or "").split("\n") if line != "" or True)
        shapes.append(
            '<p:sp><p:nvSpPr><p:cNvPr id="{id}" name="Content"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
            '<p:spPr><a:xfrm><a:off x="838200" y="1825625"/><a:ext cx="10515600" cy="4351338"/></a:xfrm></p:spPr>'
            '<p:txBody><a:bodyPr/><a:lstStyle/>{paras}</p:txBody></p:sp>'.format(
                id=next_id, paras=paras))
        next_id += 1
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:sld xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '">'
            '<p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/>' + "".join(shapes) +
            '</p:spTree></p:cSld>'
            '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'
            '</p:sld>')


def _notes_slide_xml(text: str) -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:notes xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '">'
            '<p:cSld><p:spTree>'
            '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>'
            '<p:grpSpPr/>'
            '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr>'
            '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>'
            '<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder"/>'
            '<p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr>'
            '<p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>'
            '<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>'
            '<a:p><a:r><a:rPr lang="zh-CN"/><a:t>' + _esc(text) + '</a:t></a:r></a:p>'
            '</p:txBody></p:sp>'
            '</p:spTree></p:cSld>'
            '</p:notes>')


def _esc(s) -> str:
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def _content_types_xml(slide_nums, notes_nums=()) -> str:
    overrides = [
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>',
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>',
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>',
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>',
    ]
    for n in slide_nums:
        overrides.append(
            '<Override PartName="/ppt/slides/slide{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>'.format(n=n))
    for n in notes_nums:
        overrides.append(
            '<Override PartName="/ppt/notesSlides/notesSlide{n}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>'.format(n=n))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Types xmlns="' + NS_CT + '">'
            '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
            '<Default Extension="xml" ContentType="application/xml"/>'
            + "".join(overrides) + '</Types>')


def _root_rels_xml() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="' + NS_REL + '">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>'
            '</Relationships>')


def _presentation_xml(slide_rids) -> str:
    slides = "".join('<p:sldId id="{id}" r:id="{rid}"/>'.format(id=256 + i, rid=rid)
                     for i, rid in enumerate(slide_rids))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<p:presentation xmlns:a="' + NS_A + '" xmlns:r="' + NS_R + '" xmlns:p="' + NS_P + '" '
            'saveSubsetFonts="1">'
            '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
            '<p:sldIdLst>' + slides + '</p:sldIdLst>'
            '<p:sldSz cx="12192000" cy="6858000"/>'
            '<p:notesSz cx="6858000" cy="9144000"/>'
            '</p:presentation>')


def _presentation_rels_xml(slide_nums, notes_nums=()) -> str:
    rels = ['<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>']
    rels.append('<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>')
    rid = 3
    for n in slide_nums:
        rels.append('<Relationship Id="rId{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide{n}.xml"/>'.format(rid=rid, n=n))
        rid += 1
    for n in notes_nums:
        rels.append('<Relationship Id="rId{rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>'.format(rid=rid))
        rid += 1
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="' + NS_REL + '">' + "".join(rels) + '</Relationships>')


def _master_rels_xml() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="' + NS_REL + '">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
            '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>'
            '</Relationships>')


def _layout_rels_xml() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="' + NS_REL + '">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>'
            '</Relationships>')


# ---------- PPTX 包操作核心 ----------

class LinuxPptBridge:
    """纯标准库 pptx 读写后端"""

    def __init__(self):
        self._parts = {}       # part_name -> str/bytes
        self._path = None
        self._next_rid = 100   # presentation.xml.rels 的 slide rId 计数器
        self._next_slide_num = 1
        self._cli = WpsCli()

    # ----- 包构建 -----

    def _ensure_pres(self):
        if "ppt/presentation.xml" not in self._parts:
            raise RuntimeError("没有活动演示文稿，请先 createPresentation 或 openPresentation")

    def _slide_parts(self):
        """返回已排序的 slide 编号列表"""
        nums = []
        for name in self._parts:
            m = re.match(r"^ppt/slides/slide(\d+)\.xml$", name)
            if m:
                nums.append(int(m.group(1)))
        return sorted(nums)

    def _slide_rids(self):
        """从 presentation.xml 提取 slide 的 rId 顺序列表"""
        root = ET.fromstring(self._parts["ppt/presentation.xml"])
        rids = []
        lst = root.find(_q(NS_P, "sldIdLst"))
        if lst is not None:
            for sld in lst.findall(_q(NS_P, "sldId")):
                rid = sld.get(_q(NS_R, "id"))
                rids.append(rid)
        return rids

    def _rebuild_package(self):
        """根据当前 slides 重建 presentation.xml / rels / Content_Types"""
        nums = self._slide_parts()
        # rId: master=rId1, theme=rId2, slides 从 rId3 起
        slide_rids = ["rId{}".format(3 + i) for i in range(len(nums))]
        self._parts["ppt/presentation.xml"] = _presentation_xml(slide_rids)
        self._parts["ppt/_rels/presentation.xml.rels"] = _presentation_rels_xml(nums)
        notes_nums = []
        for name in self._parts:
            m = re.match(r"^ppt/notesSlides/notesSlide(\d+)\.xml$", name)
            if m:
                notes_nums.append(int(m.group(1)))
        self._parts["[Content_Types].xml"] = _content_types_xml(nums, sorted(notes_nums))

    # ----- action 实现 -----

    def ping(self, params):
        return {"success": True, "data": {"message": "pong", "backend": "stdlib-openxml"}}

    def getAppInfo(self, params):
        return {"success": True, "data": {
            "app": "WPS演示(Linux文件模式)",
            "backend": "stdlib OpenXML",
            "wpsCli": self._cli.info(),
        }}

    def createPresentation(self, params):
        self._parts = {
            "[Content_Types].xml": _content_types_xml([]),
            "_rels/.rels": _root_rels_xml(),
            "ppt/presentation.xml": _presentation_xml([]),
            "ppt/_rels/presentation.xml.rels": _presentation_rels_xml([]),
            "ppt/slideMasters/slideMaster1.xml": _slide_master_xml(),
            "ppt/slideMasters/_rels/slideMaster1.xml.rels": _master_rels_xml(),
            "ppt/slideLayouts/slideLayout1.xml": _slide_layout_xml(),
            "ppt/slideLayouts/_rels/slideLayout1.xml.rels": _layout_rels_xml(),
            "ppt/theme/theme1.xml": _theme_xml(),
        }
        self._path = params.get("filePath")
        return {"success": True, "data": {"name": "演示文稿1", "slideCount": 0,
                                         "path": self._path or "(未保存)"}}

    def openPresentation(self, params):
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
            return {"success": False, "error": f"打开失败（仅支持 .pptx）: {e}"}
        if params.get("openInWps"):
            self._cli.open_file("ppt", fp)
        return {"success": True, "data": {"name": os.path.basename(fp),
                                         "slideCount": len(self._slide_parts()), "path": fp}}

    def save(self, params):
        self._ensure_pres()
        if not self._path:
            return {"success": False, "error": "演示文稿未指定保存路径，请用 saveAs"}
        self._write_zip(self._path)
        return {"success": True, "data": {"path": self._path}}

    def saveAs(self, params):
        self._ensure_pres()
        fp = params.get("filePath")
        if not fp:
            return {"success": False, "error": "缺少 filePath"}
        fp = os.path.abspath(fp)
        os.makedirs(os.path.dirname(fp), exist_ok=True)
        self._write_zip(fp)
        self._path = fp
        return {"success": True, "data": {"path": fp}}

    def _write_zip(self, fp):
        self._rebuild_package()
        with zipfile.ZipFile(fp, "w", zipfile.ZIP_DEFLATED) as z:
            # [Content_Types].xml 必须第一个写入
            z.writestr("[Content_Types].xml", self._parts["[Content_Types].xml"])
            for name, content in self._parts.items():
                if name == "[Content_Types].xml":
                    continue
                if isinstance(content, str):
                    content = content.encode("utf-8")
                z.writestr(name, content)

    def getSlideCount(self, params):
        self._ensure_pres()
        return {"success": True, "data": {"slideCount": len(self._slide_parts())}}

    def addSlide(self, params):
        self._ensure_pres()
        layout = params.get("layout", "title_content")
        title = params.get("title", "")
        content = params.get("content", "")
        num = self._next_slide_num
        # 避免与已有编号冲突
        while f"ppt/slides/slide{num}.xml" in self._parts:
            num += 1
        self._next_slide_num = num + 1
        self._parts[f"ppt/slides/slide{num}.xml"] = _new_slide_xml(title, content, layout)
        # 幻灯片级 rels（引用 layout）
        self._parts[f"ppt/slides/_rels/slide{num}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="' + NS_REL + '">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>'
            '</Relationships>')
        self._rebuild_package()
        index = len(self._slide_parts())
        return {"success": True, "data": {"slideIndex": index, "layout": layout}}

    def _slide_file_by_index(self, slide_index):
        """slideIndex 1-based -> part name"""
        nums = self._slide_parts()
        if slide_index < 1 or slide_index > len(nums):
            return None
        return f"ppt/slides/slide{nums[slide_index - 1]}.xml"

    def _parse_slide(self, slide_index):
        name = self._slide_file_by_index(int(slide_index))
        if not name:
            return None, None
        root = ET.fromstring(self._parts[name])
        return name, root

    def _write_slide(self, name, root):
        self._parts[name] = ET.tostring(root, encoding="unicode", xml_declaration=True)

    def deleteSlide(self, params):
        self._ensure_pres()
        idx = int(params.get("slideIndex", 0))
        name = self._slide_file_by_index(idx)
        if not name:
            return {"success": False, "error": f"未找到幻灯片 {idx}"}
        num = re.match(r"ppt/slides/slide(\d+)\.xml$", name).group(1)
        del self._parts[name]
        rels = f"ppt/slides/_rels/slide{num}.xml.rels"
        if rels in self._parts:
            del self._parts[rels]
        self._rebuild_package()
        return {"success": True, "data": {"slideCount": len(self._slide_parts())}}

    def _find_ph_sp(self, root, ph_type):
        """按占位符类型查找 shape 元素"""
        for sp in root.iter(_q(NS_P, "sp")):
            nv = sp.find(_q(NS_P, "nvSpPr"))
            if nv is None:
                continue
            nvpr = nv.find(_q(NS_P, "nvPr"))
            if nvpr is None:
                continue
            ph = nvpr.find(_q(NS_P, "ph"))
            if ph is not None and ph.get("type", "body") == ph_type:
                return sp
        return None

    @staticmethod
    def _set_sp_text(sp, text):
        """设置 shape 内文本（替换所有段落为单一 runs）"""
        tx = sp.find(_q(NS_P, "txBody"))
        if tx is None:
            return
        # 清掉已有 a:p
        for p in tx.findall(_q(NS_A, "p")):
            tx.remove(p)
        lines = str(text).split("\n")
        for line in lines:
            p = ET.SubElement(tx, _q(NS_A, "p"))
            r = ET.SubElement(p, _q(NS_A, "r"))
            rpr = ET.SubElement(r, _q(NS_A, "rPr"))
            rpr.set("lang", "zh-CN")
            t = ET.SubElement(r, _q(NS_A, "t"))
            t.text = line

    def setSlideTitle(self, params):
        self._ensure_pres()
        idx = params.get("slideIndex", 1)
        name, root = self._parse_slide(idx)
        if not root:
            return {"success": False, "error": f"未找到幻灯片 {idx}"}
        sp = self._find_ph_sp(root, "title")
        if sp is None:
            return {"success": False, "error": "该幻灯片无标题占位符"}
        self._set_sp_text(sp, params.get("title", ""))
        self._write_slide(name, root)
        return {"success": True}

    def getSlideTitle(self, params):
        self._ensure_pres()
        idx = params.get("slideIndex", 1)
        name, root = self._parse_slide(idx)
        if not root:
            return {"success": False, "error": f"未找到幻灯片 {idx}"}
        sp = self._find_ph_sp(root, "title")
        if sp is None:
            return {"success": True, "data": {"title": ""}}
        texts = [t.text or "" for t in sp.iter(_q(NS_A, "t"))]
        return {"success": True, "data": {"title": "".join(texts)}}

    def setSlideContent(self, params):
        self._ensure_pres()
        idx = params.get("slideIndex", 1)
        name, root = self._parse_slide(idx)
        if not root:
            return {"success": False, "error": f"未找到幻灯片 {idx}"}
        sp = self._find_ph_sp(root, "body")
        if sp is None:
            return {"success": False, "error": "该幻灯片无内容占位符"}
        self._set_sp_text(sp, params.get("content", ""))
        self._write_slide(name, root)
        return {"success": True}

    def getSlideInfo(self, params):
        self._ensure_pres()
        idx = params.get("slideIndex", 1)
        name, root = self._parse_slide(idx)
        if not root:
            return {"success": False, "error": f"未找到幻灯片 {idx}"}
        shapes = []
        sid = 0
        for sp in root.iter(_q(NS_P, "sp")):
            sid += 1
            nv = sp.find(_q(NS_P, "nvSpPr"))
            nm = ""
            ph_type = None
            if nv is not None:
                cnv = nv.find(_q(NS_P, "cNvPr"))
                nm = cnv.get("name", "") if cnv is not None else ""
                nvpr = nv.find(_q(NS_P, "nvPr"))
                if nvpr is not None:
                    ph = nvpr.find(_q(NS_P, "ph"))
                    if ph is not None:
                        ph_type = ph.get("type", "body")
            text = "".join(t.text or "" for t in sp.iter(_q(NS_A, "t")))
            shapes.append({"index": sid, "name": nm, "placeholder": ph_type, "text": text})
        return {"success": True, "data": {"slideIndex": idx, "shapeCount": len(shapes),
                                         "shapes": shapes}}

    def setSlideNotes(self, params):
        self._ensure_pres()
        idx = int(params.get("slideIndex", 1))
        slide_name = self._slide_file_by_index(idx)
        if not slide_name:
            return {"success": False, "error": f"未找到幻灯片 {idx}"}
        num = re.match(r"ppt/slides/slide(\d+)\.xml$", slide_name).group(1)
        self._parts[f"ppt/notesSlides/notesSlide{num}.xml"] = _notes_slide_xml(
            params.get("text", params.get("notes", "")))
        # notesSlide rels 指向 slide
        self._parts[f"ppt/notesSlides/_rels/notesSlide{num}.xml.rels"] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="' + NS_REL + '">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide' + num + '.xml"/>'
            '</Relationships>')
        # slide rels 增加 notesSlide 引用
        slide_rels_name = f"ppt/slides/_rels/slide{num}.xml.rels"
        if slide_rels_name in self._parts:
            content = self._parts[slide_rels_name] if isinstance(self._parts[slide_rels_name], str) \
                else self._parts[slide_rels_name].decode("utf-8")
            if "notesSlide" not in content:
                content = content.replace(
                    "</Relationships>",
                    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide' + num + '.xml"/></Relationships>')
                self._parts[slide_rels_name] = content
        return {"success": True}

    def getSlideNotes(self, params):
        self._ensure_pres()
        idx = int(params.get("slideIndex", 1))
        slide_name = self._slide_file_by_index(idx)
        if not slide_name:
            return {"success": False, "error": f"未找到幻灯片 {idx}"}
        num = re.match(r"ppt/slides/slide(\d+)\.xml$", slide_name).group(1)
        notes_name = f"ppt/notesSlides/notesSlide{num}.xml"
        if notes_name not in self._parts:
            return {"success": True, "data": {"notes": ""}}
        content = self._parts[notes_name]
        if isinstance(content, bytes):
            content = content.decode("utf-8")
        texts = re.findall(r"<a:t>([^<]*)</a:t>", content)
        return {"success": True, "data": {"notes": "".join(texts)}}

    def getOpenPresentations(self, params):
        if self._path:
            return {"success": True, "data": {"presentations": [os.path.basename(self._path)],
                                             "count": 1}}
        return {"success": True, "data": {"presentations": [], "count": 0}}

    def findPptText(self, params):
        self._ensure_pres()
        find_t = params.get("find_text", "")
        results = []
        for i in range(1, len(self._slide_parts()) + 1):
            _, root = self._parse_slide(i)
            if root is None:
                continue
            all_text = "".join(t.text or "" for t in root.iter(_q(NS_A, "t")))
            if find_t in all_text:
                results.append({"slideIndex": i})
        return {"success": True, "data": {"matches": results, "count": len(results)}}

    def replacePptText(self, params):
        self._ensure_pres()
        find_t = params.get("find_text", "")
        rep_t = params.get("replace_text", "")
        count = 0
        for i in range(1, len(self._slide_parts()) + 1):
            name, root = self._parse_slide(i)
            if root is None:
                continue
            changed = False
            for t in root.iter(_q(NS_A, "t")):
                if t.text and find_t in t.text:
                    t.text = t.text.replace(find_t, rep_t)
                    count += 1
                    changed = True
            if changed:
                self._write_slide(name, root)
        return {"success": True, "data": {"replaced": count}}

    def setSlideSize(self, params):
        self._ensure_pres()
        w = int(params.get("width", 12192000))
        h = int(params.get("height", 6858000))
        root = ET.fromstring(self._parts["ppt/presentation.xml"])
        sz = root.find(_q(NS_P, "sldSz"))
        if sz is not None:
            sz.set("cx", str(w))
            sz.set("cy", str(h))
        self._parts["ppt/presentation.xml"] = ET.tostring(root, encoding="unicode",
                                                          xml_declaration=True)
        return {"success": True}

    def convertToPDF(self, params):
        fp = params.get("filePath") or self._path
        if not fp or not os.path.isfile(fp):
            return {"success": False, "error": "缺少 filePath 或文件不存在"}
        if self._path == fp and "ppt/presentation.xml" in self._parts:
            self.save({})  # 先保存最新内容
        return self._cli.convert_to_pdf("ppt", fp, params.get("outputPath"))

    def closePresentation(self, params):
        if params.get("save") and self._path:
            self.save({})
        self._parts = {}
        self._path = None
        return {"success": True}

    # ---------- 统一分发 ----------

    def execute(self, action: str, params: dict) -> dict:
        params = params or {}
        if action == "ping":
            return self.ping(params)
        fn = getattr(self, action, None)
        if fn is None or not callable(fn):
            return {"success": False,
                    "error": f"Linux 文件模式未实现 action '{action}'（动画/3D/切换等高级特性需 WPS GUI）"}
        try:
            return fn(params)
        except RuntimeError as e:
            return {"success": False, "error": str(e)}
        except Exception as e:
            return {"success": False, "error": f"{type(e).__name__}: {e}"}
