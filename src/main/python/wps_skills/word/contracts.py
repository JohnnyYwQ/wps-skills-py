"""Authoritative Word Application Contract Set and derived Action Index."""

import re
import unicodedata

from wps_skills.core.action_session import ActionContract, ApplicationContractSet


def _host_absolute_path(value, suffixes):
    if any(
        unicodedata.category(character) in {"Cc", "Cs"}
        for character in value
    ):
        return False
    drive_absolute = re.match(r"^[A-Za-z]:[\\/]", value) is not None
    unc_absolute = re.match(
        r"^\\\\[^\\/]+[\\/][^\\/]+[\\/].+$",
        value,
    ) is not None
    posix_absolute = value.startswith("/") and not value.startswith("//")
    if not (drive_absolute or unc_absolute or posix_absolute):
        return False
    if (drive_absolute or unc_absolute) and not _windows_path_is_valid(value):
        return False
    return value.casefold().endswith(suffixes)


def _windows_path_is_valid(value):
    tail = value[3:] if re.match(r"^[A-Za-z]:[\\/]", value) else value[2:]
    components = re.split(r"[\\/]", tail)
    if not components or any(not component for component in components):
        return False
    reserved = re.compile(
        r"^(?:CON|PRN|AUX|NUL|COM[1-9]|LPT[1-9])(?:\..*)?$",
        re.IGNORECASE,
    )
    return all(
        re.search(r'[<>:"|?*]', component) is None
        and not component.endswith((" ", "."))
        and reserved.fullmatch(component) is None
        for component in components
    )


def _word_run_text(value):
    return bool(value) and all(
        unicodedata.category(character) not in {"Cc", "Cs"}
        and character not in {"\u2028", "\u2029", "\f", "\ufffc"}
        for character in value
    )


def _word_story_text(value):
    return all(
        character in {"\t", "\n"}
        or (
            unicodedata.category(character) not in {"Cc", "Cs"}
            and character not in {"\u2028", "\u2029", "\ufffc"}
        )
        for character in value
    ) and "\r" not in value


def _word_normalized_text(value):
    return all(
        character in {"\t", "\n", "\f", "\u2028", "\ufffc"}
        or (
            unicodedata.category(character) not in {"Cc", "Cs"}
            and character != "\u2029"
        )
        for character in value
    ) and "\r" not in value


def _word_query_text(value):
    forbidden = {
        "\t", "\n", "\r", "\u2028", "\u2029", "\f", "\ufffc"
    }
    return bool(value) and not any(
        character in forbidden
        or unicodedata.category(character) in {"Cc", "Cs"}
        for character in value
    )


WORD_FORMAT_VALIDATORS = {
    "absoluteDocxPath": lambda value: _host_absolute_path(
        value,
        (".docx",),
    ),
    "absolutePdfPath": lambda value: _host_absolute_path(
        value,
        (".pdf",),
    ),
    "absoluteImagePath": lambda value: _host_absolute_path(
        value,
        (".png", ".jpg", ".jpeg"),
    ),
    "wordRunText": _word_run_text,
    "wordStoryText": _word_story_text,
    "wordNormalizedText": _word_normalized_text,
    "wordQueryText": _word_query_text,
}


def _object(properties, required=(), **keywords):
    schema = {
        "type": "object",
        "properties": properties,
        "required": tuple(required),
        "additionalProperties": False,
    }
    schema.update(keywords)
    return schema


def _array(items, *, minimum=0, maximum=None, **keywords):
    schema = {
        "type": "array",
        "items": items,
        "minItems": minimum,
    }
    if maximum is not None:
        schema["maxItems"] = maximum
    schema.update(keywords)
    return schema


def _string(*, minimum=0, maximum=None, enum=None, **keywords):
    schema = {"type": "string", "minLength": minimum}
    if maximum is not None:
        schema["maxLength"] = maximum
    if enum is not None:
        schema["enum"] = tuple(enum)
    schema.update(keywords)
    return schema


def _integer(*, minimum=None, maximum=None, **keywords):
    schema = {"type": "integer"}
    if minimum is not None:
        schema["minimum"] = minimum
    if maximum is not None:
        schema["maximum"] = maximum
    schema.update(keywords)
    return schema


def _number(*, minimum=None, maximum=None, **keywords):
    schema = {"type": "number"}
    if minimum is not None:
        schema["minimum"] = minimum
    if maximum is not None:
        schema["maximum"] = maximum
    schema.update(keywords)
    return schema


def _const(value):
    return {"const": value}


def _nullable(schema):
    return {"oneOf": (schema, {"type": "null"})}


def _length_between_points(minimum_pt, maximum_pt):
    units_per_point = {
        "pt": 1,
        "in": 1 / 72,
        "cm": 2.54 / 72,
        "mm": 25.4 / 72,
    }
    return {
        "oneOf": tuple(
            _object(
                {
                    "value": _number(
                        minimum=minimum_pt * factor,
                        maximum=maximum_pt * factor,
                    ),
                    "unit": _const(unit),
                },
                ("value", "unit"),
            )
            for unit, factor in units_per_point.items()
        )
    }


CONTENT_REVISION = _string(minimum=1, maximum=128)

CONTENT_RANGE = _object(
    {
        "start": _integer(minimum=0),
        "end": _integer(minimum=0),
        "revision": CONTENT_REVISION,
    },
    ("start", "end", "revision"),
    **{"x-ordered": ("start", "end")},
)

NONEMPTY_CONTENT_RANGE = _object(
    {
        "start": _integer(minimum=0),
        "end": _integer(minimum=0),
        "revision": CONTENT_REVISION,
    },
    ("start", "end", "revision"),
    **{"x-strictOrdered": ("start", "end")},
)

BODY_ANCHOR = {
    "oneOf": (
        _object({"kind": _const("documentStart")}, ("kind",)),
        _object({"kind": _const("documentEnd")}, ("kind",)),
        _object(
            {"kind": _const("before"), "range": CONTENT_RANGE},
            ("kind", "range"),
        ),
        _object(
            {"kind": _const("after"), "range": CONTENT_RANGE},
            ("kind", "range"),
        ),
    ),
}

BODY_SCOPE = {
    "oneOf": (
        _object({"kind": _const("document")}, ("kind",)),
        _object(
            {"kind": _const("range"), "range": CONTENT_RANGE},
            ("kind", "range"),
        ),
    ),
}

POSITIVE_LENGTH = _length_between_points(0.01, 1584)
OFFSET_LENGTH = _length_between_points(-1584, 1584)
MARGIN_LENGTH = _length_between_points(0, 720)

POINT_LENGTH = _object(
    {
        "value": _number(minimum=0, maximum=1584),
        "unit": _const("pt"),
    },
    ("value", "unit"),
)

POSITIVE_POINT_LENGTH = _object(
    {
        "value": _number(minimum=0.01, maximum=1584),
        "unit": _const("pt"),
    },
    ("value", "unit"),
)

POINT_OFFSET = _object(
    {
        "value": _number(minimum=-1584, maximum=1584),
        "unit": _const("pt"),
    },
    ("value", "unit"),
)

DOCUMENT_STATE = _object(
    {
        "persistenceState": _string(
            enum=("unsaved", "saved", "modified")
        ),
        "readOnly": {"type": "boolean"},
    },
    ("persistenceState", "readOnly"),
)

CREATED_DOCUMENT_STATE = _object(
    {
        "persistenceState": _const("unsaved"),
        "readOnly": _const(False),
    },
    ("persistenceState", "readOnly"),
)

SAVED_DOCUMENT_STATE = _object(
    {
        "persistenceState": _const("saved"),
        "readOnly": {"type": "boolean"},
    },
    ("persistenceState", "readOnly"),
)

OPENED_DOCUMENT_STATE = _object(
    {
        "persistenceState": _string(enum=("saved", "modified")),
        "readOnly": {"type": "boolean"},
    },
    ("persistenceState", "readOnly"),
)

DOCX_PATH = _string(
    minimum=1,
    maximum=4096,
    format="absoluteDocxPath",
)

PDF_PATH = _string(
    minimum=1,
    maximum=4096,
    format="absolutePdfPath",
)

IMAGE_PATH = _string(
    minimum=1,
    maximum=4096,
    format="absoluteImagePath",
)

DOCX_ARTIFACT = _object(
    {
        "path": DOCX_PATH,
        "format": _const("docx"),
        "sizeBytes": _integer(minimum=1),
    },
    ("path", "format", "sizeBytes"),
)

PDF_ARTIFACT = _object(
    {
        "path": PDF_PATH,
        "format": _const("pdf"),
        "sizeBytes": _integer(minimum=1),
    },
    ("path", "format", "sizeBytes"),
)

TEXT_FORMAT_PATCH = _object(
    {
        "fontFamily": _string(minimum=1, maximum=128),
        "westernFontFamily": _string(minimum=1, maximum=128),
        "eastAsiaFontFamily": _string(minimum=1, maximum=128),
        "fontSizePt": _number(minimum=1, maximum=300),
        "bold": {"type": "boolean"},
        "italic": {"type": "boolean"},
        "underline": _string(enum=("none", "single")),
        "color": _string(pattern=r"^#[0-9A-Fa-f]{6}$"),
    },
    (),
    **{
        "x-atLeastOne": (
            "fontFamily",
            "westernFontFamily",
            "eastAsiaFontFamily",
            "fontSizePt",
            "bold",
            "italic",
            "underline",
            "color",
        )
    },
)

LINE_SPACING = {
    "oneOf": (
        _object({"kind": _const("single")}, ("kind",)),
        _object({"kind": _const("oneAndHalf")}, ("kind",)),
        _object({"kind": _const("double")}, ("kind",)),
        _object(
            {
                "kind": _const("exact"),
                "points": _number(minimum=0.01, maximum=1584),
            },
            ("kind", "points"),
        ),
        _object(
            {
                "kind": _const("atLeast"),
                "points": _number(minimum=0.01, maximum=1584),
            },
            ("kind", "points"),
        ),
        _object(
            {
                "kind": _const("multiple"),
                "value": _number(minimum=0.01, maximum=100),
            },
            ("kind", "value"),
        ),
    ),
}

PARAGRAPH_FORMAT_PATCH = _object(
    {
        "alignment": _string(
            enum=("left", "center", "right", "justify")
        ),
        "lineSpacing": LINE_SPACING,
        "spaceBeforePt": _number(minimum=0, maximum=1584),
        "spaceAfterPt": _number(minimum=0, maximum=1584),
        "leftIndentPt": _number(minimum=0, maximum=1584),
        "rightIndentPt": _number(minimum=0, maximum=1584),
        "firstLineIndentPt": _number(minimum=-1584, maximum=1584),
    },
    (),
    **{
        "x-atLeastOne": (
            "alignment",
            "lineSpacing",
            "spaceBeforePt",
            "spaceAfterPt",
            "leftIndentPt",
            "rightIndentPt",
            "firstLineIndentPt",
        )
    },
)

TEXT_RUN = _object(
    {
        "text": _string(
            minimum=1,
            maximum=32768,
            format="wordRunText",
            **{"x-maxUtf16Length": 32768},
        ),
        "format": TEXT_FORMAT_PATCH,
    },
    ("text",),
)

TEXT_RUNS = _array(
    TEXT_RUN,
    minimum=1,
    maximum=128,
    **{"x-maxUtf16Text": 262144},
)

INLINE_TEXT_BLOCK = _object(
    {
        "kind": _const("text"),
        "runs": TEXT_RUNS,
    },
    ("kind", "runs"),
)

PARAGRAPH_BLOCK = _object(
    {
        "kind": _const("paragraph"),
        "runs": _array(
            TEXT_RUN,
            minimum=0,
            maximum=128,
            **{"x-maxUtf16Text": 262144},
        ),
        "format": PARAGRAPH_FORMAT_PATCH,
    },
    ("kind", "runs"),
)

HEADING_BLOCK = _object(
    {
        "kind": _const("heading"),
        "level": _integer(minimum=1, maximum=9),
        "runs": TEXT_RUNS,
        "format": PARAGRAPH_FORMAT_PATCH,
    },
    ("kind", "level", "runs"),
)

STRUCTURED_BLOCKS = {
    "oneOf": (
        _array(
            INLINE_TEXT_BLOCK,
            minimum=1,
            maximum=1,
            **{"x-maxUtf16Text": 262144},
        ),
        _array(
            {"oneOf": (PARAGRAPH_BLOCK, HEADING_BLOCK)},
            minimum=1,
            maximum=128,
            **{"x-maxUtf16Text": 262144},
        ),
    ),
}

TEXT_QUERY = _object(
    {
        "scope": BODY_SCOPE,
        "text": _string(
            minimum=1,
            maximum=4096,
            format="wordQueryText",
        ),
        "caseSensitive": {"type": "boolean"},
        "wholeWord": {"type": "boolean"},
    },
    ("scope", "text", "caseSensitive", "wholeWord"),
)

EFFECTIVE_TEXT_FORMAT = _object(
    {
        "fontFamily": _nullable(_string()),
        "westernFontFamily": _nullable(_string()),
        "eastAsiaFontFamily": _nullable(_string()),
        "fontSizePt": _nullable(_number(minimum=0)),
        "bold": _nullable({"type": "boolean"}),
        "italic": _nullable({"type": "boolean"}),
        "underline": _nullable(
            _string(enum=("none", "single", "other"))
        ),
        "color": _nullable(
            {
                "oneOf": (
                    _string(pattern=r"^#[0-9A-F]{6}$"),
                    _const("automatic"),
                )
            }
        ),
    },
    (
        "fontFamily",
        "westernFontFamily",
        "eastAsiaFontFamily",
        "fontSizePt",
        "bold",
        "italic",
        "underline",
        "color",
    ),
)

OBSERVED_RUN = _object(
    {
        "range": CONTENT_RANGE,
        "text": _string(maximum=32768, format="wordNormalizedText"),
        "format": EFFECTIVE_TEXT_FORMAT,
    },
    ("range", "text", "format"),
)

OBSERVED_LINE_SPACING = _nullable({
    "oneOf": LINE_SPACING["oneOf"] + (
        _object({"kind": _const("other")}, ("kind",)),
    )
})

EFFECTIVE_PARAGRAPH_FORMAT = _object(
    {
        "alignment": _nullable(
            _string(enum=("left", "center", "right", "justify", "other"))
        ),
        "lineSpacing": OBSERVED_LINE_SPACING,
        "spaceBeforePt": _nullable(_number()),
        "spaceAfterPt": _nullable(_number()),
        "leftIndentPt": _nullable(_number()),
        "rightIndentPt": _nullable(_number()),
        "firstLineIndentPt": _nullable(_number()),
    },
    (
        "alignment",
        "lineSpacing",
        "spaceBeforePt",
        "spaceAfterPt",
        "leftIndentPt",
        "rightIndentPt",
        "firstLineIndentPt",
    ),
)

PARAGRAPH_SNAPSHOT_BASE = {
    "range": CONTENT_RANGE,
    "complete": {"type": "boolean"},
    "text": _string(maximum=65536, format="wordNormalizedText"),
    "runs": _array(OBSERVED_RUN, maximum=2048),
    "format": EFFECTIVE_PARAGRAPH_FORMAT,
}

PARAGRAPH_SNAPSHOT = {
    "oneOf": (
        _object(
            {
                **PARAGRAPH_SNAPSHOT_BASE,
                "kind": _const("paragraph"),
            },
            (
                "kind",
                "range",
                "complete",
                "text",
                "runs",
                "format",
            ),
        ),
        _object(
            {
                **PARAGRAPH_SNAPSHOT_BASE,
                "kind": _const("heading"),
                "level": _integer(minimum=1, maximum=9),
            },
            (
                "kind",
                "level",
                "range",
                "complete",
                "text",
                "runs",
                "format",
            ),
        ),
    )
}

SECTION_SELECTOR = {
    "oneOf": (
        _object(
            {
                "kind": _const("all"),
                "revision": CONTENT_REVISION,
            },
            ("kind", "revision"),
        ),
        _object(
            {
                "kind": _const("indexes"),
                "indexes": _array(
                    _integer(minimum=0),
                    minimum=1,
                    maximum=64,
                    uniqueItems=True,
                    **{"x-strictlyIncreasing": True},
                ),
                "revision": CONTENT_REVISION,
            },
            ("kind", "indexes", "revision"),
        ),
    )
}

MARGINS_INPUT = _object(
    {
        "top": MARGIN_LENGTH,
        "right": MARGIN_LENGTH,
        "bottom": MARGIN_LENGTH,
        "left": MARGIN_LENGTH,
    },
    ("top", "right", "bottom", "left"),
)

MARGINS_SNAPSHOT = _object(
    {
        "top": POINT_LENGTH,
        "right": POINT_LENGTH,
        "bottom": POINT_LENGTH,
        "left": POINT_LENGTH,
    },
    ("top", "right", "bottom", "left"),
)

LAYOUT_SNAPSHOT = _object(
    {
        "orientation": _string(enum=("portrait", "landscape")),
        "margins": MARGINS_SNAPSHOT,
    },
    ("orientation", "margins"),
)

HEADER_FOOTER_STORY = _object(
    {
        "area": _string(enum=("header", "footer")),
        "variant": _string(
            enum=("primary", "firstPage", "evenPages")
        ),
        "exists": {"type": "boolean"},
        "linkToPrevious": {"type": "boolean"},
        "text": _string(
            maximum=32768,
            format="wordStoryText",
            **{"x-maxUtf16Length": 32768},
        ),
    },
    ("area", "variant", "exists", "linkToPrevious", "text"),
)

HEADER_FOOTER_SNAPSHOT = _object(
    {
        "firstPageEnabled": {"type": "boolean"},
        "evenPagesEnabled": {"type": "boolean"},
        "stories": _array(
            HEADER_FOOTER_STORY,
            minimum=6,
            maximum=6,
            **{"x-maxUtf16Text": 196608},
        ),
    },
    ("firstPageEnabled", "evenPagesEnabled", "stories"),
)

SECTION_SNAPSHOT = _object(
    {
        "index": _integer(minimum=0),
        "layout": LAYOUT_SNAPSHOT,
        "headerFooter": HEADER_FOOTER_SNAPSHOT,
    },
    ("index", "layout", "headerFooter"),
)

STRUCTURE_SNAPSHOT = _object(
    {
        "paragraphCount": _integer(minimum=0),
        "headingCount": _integer(minimum=0),
        "tableCount": _integer(minimum=0),
        "inlineImageCount": _integer(minimum=0),
        "floatingImageCount": _integer(minimum=0),
        "sectionCount": _integer(minimum=1, maximum=64),
        "pageBreakCount": _integer(minimum=0),
        "sectionBreakCount": _integer(minimum=0),
        "sections": _array(SECTION_SNAPSHOT, minimum=1, maximum=64),
    },
    (
        "paragraphCount",
        "headingCount",
        "tableCount",
        "inlineImageCount",
        "floatingImageCount",
        "sectionCount",
        "pageBreakCount",
        "sectionBreakCount",
        "sections",
    ),
)

TABLE_DATA = _array(
    _array(
        _string(
            maximum=32768,
            format="wordStoryText",
            **{"x-maxUtf16Length": 32768},
        ),
        minimum=1,
        maximum=50,
    ),
    minimum=1,
    maximum=100,
    **{
        "x-rectangular": True,
        "x-maxCells": 2000,
        "x-maxUtf16Length": 1048576,
    },
)

ALTERNATIVE_TEXT = {
    "oneOf": (
        _object({"kind": _const("decorative")}, ("kind",)),
        _object(
            {
                "kind": _const("description"),
                "text": _string(
                    minimum=1,
                    maximum=2048,
                    format="wordStoryText",
                    **{"x-maxUtf16Length": 2048},
                ),
            },
            ("kind", "text"),
        ),
    )
}

IMAGE_PLACEMENT = {
    "oneOf": (
        _object({"kind": _const("inline")}, ("kind",)),
        _object(
            {
                "kind": _const("floating"),
                "wrap": _string(
                    enum=(
                        "square",
                        "topBottom",
                        "behindText",
                        "inFrontOfText",
                    )
                ),
                "horizontal": _object(
                    {
                        "relativeTo": _string(
                            enum=("page", "margin", "column")
                        ),
                        "offset": OFFSET_LENGTH,
                    },
                    ("relativeTo", "offset"),
                ),
                "vertical": _object(
                    {
                        "relativeTo": _string(
                            enum=("page", "margin", "paragraph")
                        ),
                        "offset": OFFSET_LENGTH,
                    },
                    ("relativeTo", "offset"),
                ),
            },
            ("kind", "wrap", "horizontal", "vertical"),
        ),
    )
}

IMAGE_SIZE = {
    "oneOf": (
        _object({"kind": _const("intrinsic")}, ("kind",)),
        _object(
            {"kind": _const("width"), "width": POSITIVE_LENGTH},
            ("kind", "width"),
        ),
        _object(
            {"kind": _const("height"), "height": POSITIVE_LENGTH},
            ("kind", "height"),
        ),
        _object(
            {
                "kind": _const("box"),
                "width": POSITIVE_LENGTH,
                "height": POSITIVE_LENGTH,
                "fit": _string(enum=("contain", "stretch")),
            },
            ("kind", "width", "height", "fit"),
        ),
    )
}

OBSERVED_IMAGE_SOURCE = _object(
    {
        "mediaType": _string(enum=("image/png", "image/jpeg")),
        "byteLength": _integer(minimum=1, maximum=26214400),
        "sha256": _string(pattern=r"^[0-9a-f]{64}$"),
    },
    ("mediaType", "byteLength", "sha256"),
)

OBSERVED_IMAGE_SIZE = _object(
    {
        "width": POSITIVE_POINT_LENGTH,
        "height": POSITIVE_POINT_LENGTH,
    },
    ("width", "height"),
)

OBSERVED_IMAGE = {
    "oneOf": (
        _object(
            {
                "kind": _const("inline"),
                "range": NONEMPTY_CONTENT_RANGE,
                "embedded": _const(True),
                "source": OBSERVED_IMAGE_SOURCE,
                "size": OBSERVED_IMAGE_SIZE,
                "alternativeText": ALTERNATIVE_TEXT,
            },
            (
                "kind",
                "range",
                "embedded",
                "source",
                "size",
                "alternativeText",
            ),
        ),
        _object(
            {
                "kind": _const("floating"),
                "anchorRange": CONTENT_RANGE,
                "embedded": _const(True),
                "source": OBSERVED_IMAGE_SOURCE,
                "size": OBSERVED_IMAGE_SIZE,
                "wrap": _string(
                    enum=(
                        "square",
                        "topBottom",
                        "behindText",
                        "inFrontOfText",
                    )
                ),
                "horizontal": _object(
                    {
                        "relativeTo": _string(
                            enum=("page", "margin", "column")
                        ),
                        "offset": POINT_OFFSET,
                    },
                    ("relativeTo", "offset"),
                ),
                "vertical": _object(
                    {
                        "relativeTo": _string(
                            enum=("page", "margin", "paragraph")
                        ),
                        "offset": POINT_OFFSET,
                    },
                    ("relativeTo", "offset"),
                ),
                "alternativeText": ALTERNATIVE_TEXT,
            },
            (
                "kind",
                "anchorRange",
                "embedded",
                "source",
                "size",
                "wrap",
                "horizontal",
                "vertical",
                "alternativeText",
            ),
        ),
    )
}

HEADER_FOOTER_OPERATION = {
    "oneOf": (
        _object(
            {
                "kind": _const("replace"),
                "text": _string(
                    minimum=1,
                    maximum=32768,
                    format="wordStoryText",
                    **{"x-maxUtf16Length": 32768},
                ),
            },
            ("kind", "text"),
        ),
        _object({"kind": _const("clear")}, ("kind",)),
        _object({"kind": _const("linkToPrevious")}, ("kind",)),
    )
}

HEADER_FOOTER_UPDATE = _object(
    {
        "area": _string(enum=("header", "footer")),
        "variant": _string(
            enum=("primary", "firstPage", "evenPages")
        ),
        "operation": HEADER_FOOTER_OPERATION,
    },
    ("area", "variant", "operation"),
)

OBSERVED_HEADER_FOOTER_STORY = _object(
    {
        "sectionIndex": _integer(minimum=0),
        "area": _string(enum=("header", "footer")),
        "variant": _string(
            enum=("primary", "firstPage", "evenPages")
        ),
        "variantEnabled": {"type": "boolean"},
        "exists": {"type": "boolean"},
        "linkToPrevious": {"type": "boolean"},
        "text": _string(
            maximum=32768,
            format="wordStoryText",
            **{"x-maxUtf16Length": 32768},
        ),
    },
    (
        "sectionIndex",
        "area",
        "variant",
        "variantEnabled",
        "exists",
        "linkToPrevious",
        "text",
    ),
)

OVERWRITE_POLICY = _string(
    enum=("failIfExists", "replaceExisting")
)

BINDING_ERRORS = (
    "DOCUMENT_CLOSED",
    "DOCUMENT_BINDING_UNAVAILABLE",
)

MUTATION_ERRORS = BINDING_ERRORS + (
    "DOCUMENT_READ_ONLY",
)

CONTENT_READ_ERRORS = BINDING_ERRORS + (
    "STALE_CONTENT_RANGE",
    "CONTENT_RANGE_OUT_OF_BOUNDS",
    "CONTENT_LIMIT_EXCEEDED",
    "CONTENT_CHANGED_DURING_ACTION",
)

CONTENT_MUTATION_ERRORS = CONTENT_READ_ERRORS + (
    "DOCUMENT_READ_ONLY",
    "CONTENT_PROTECTED",
)

COMMON_CONTRACT_ERRORS = (
    "INVALID_PARAMS",
    "INVALID_RESULT",
    "RESPONSE_LOST",
    "WORD_CAPABILITY_UNAVAILABLE",
)

_WORD_EXAMPLE_PARAMS = {
    "createDocument": {},
    "openDocument": {"path": "C:/docs/report.docx"},
    "writeContent": {
        "anchor": {"kind": "documentEnd"},
        "blocks": ({
            "kind": "paragraph",
            "runs": ({
                "text": "你好, world",
                "format": {
                    "westernFontFamily": "Times New Roman",
                    "eastAsiaFontFamily": "Microsoft YaHei",
                },
            },),
        },),
    },
    "inspectDocument": {
        "scope": {"kind": "document"},
        "limits": {
            "maxTextCharacters": 4096,
            "maxParagraphs": 128,
            "maxRuns": 512,
        },
    },
    "findContent": {
        "query": {
            "scope": {"kind": "document"},
            "text": "Hello",
            "caseSensitive": False,
            "wholeWord": True,
        },
        "limit": 50,
    },
    "replaceContent": {
        "target": {
            "kind": "query",
            "query": {
                "scope": {"kind": "document"},
                "text": "old",
                "caseSensitive": True,
                "wholeWord": True,
            },
            "expectedMatchCount": 1,
        },
        "replacement": {
            "kind": "text",
            "runs": ({"text": "new"},),
        },
    },
    "insertTable": {
        "anchor": {"kind": "documentEnd"},
        "data": (("Name", "Value"), ("A", "1")),
        "headerRow": True,
    },
    "insertImage": {
        "anchor": {"kind": "documentEnd"},
        "source": {"kind": "file", "path": "C:/images/logo.png"},
        "placement": {"kind": "inline"},
        "size": {"kind": "intrinsic"},
        "alternativeText": {"kind": "decorative"},
    },
    "setHeaderFooter": {
        "sections": {"kind": "all", "revision": "revision-1"},
        "updates": ({
            "area": "header",
            "variant": "primary",
            "operation": {"kind": "replace", "text": "Report"},
        },),
    },
    "setPageLayout": {
        "sections": {"kind": "all", "revision": "revision-1"},
        "layout": {"orientation": "landscape"},
    },
    "insertBreak": {
        "anchor": {"kind": "documentEnd"},
        "type": "page",
    },
    "save": {},
    "saveAs": {
        "outputPath": "C:/docs/output.docx",
        "overwritePolicy": "failIfExists",
    },
    "exportPdf": {
        "outputPath": "C:/docs/output.pdf",
        "overwritePolicy": "failIfExists",
    },
}

_WORD_CONSTRAINTS = {
    "createDocument": (
        "Creates one blank unsaved document and never accepts a path.",
        "Binding and Lease commit before a successful response is observable.",
    ),
    "openDocument": (
        "The path is an absolute existing .docx resolved by stable identity.",
        "Active UI state, display name, and open order never select the document.",
    ),
    "writeContent": (
        "Blocks are either one inline text block or a paragraph/heading sequence.",
        "fontFamily is a unified setting and cannot be mixed in one run with westernFontFamily or eastAsiaFontFamily.",
        "All limits, revision, bounds, alignment, and protection checks precede writing.",
    ),
    "inspectDocument": (
        "Every returned fact belongs to one coherent Content Revision.",
        "Body output may be bounded, but section/layout/header-footer facts never truncate.",
    ),
    "findContent": (
        "Search is literal, bounded, non-wrapping, and non-overlapping.",
        "All matches and remaining range belong to one Content Revision.",
    ),
    "replaceContent": (
        "All targets and expected match counts are preflighted before the first write.",
        "Multi-match replacement executes from document end toward start and is never replayed.",
    ),
    "insertTable": (
        "Data is a non-empty rectangular plain-text matrix within hard ceilings.",
        "Success uses read-back cell data rather than request echo.",
    ),
    "insertImage": (
        "Only staged embedded PNG/JPEG files are accepted; URLs and links are forbidden.",
        "Success reads back media hash, dimensions, placement, anchor, and alternative text.",
    ),
    "setHeaderFooter": (
        "Section selection is revision-bound and update area/variant pairs are unique.",
        "linkToPrevious never targets section 0; selecting all sections therefore forbids it.",
        "All selected stories are preflighted before any write and read back after mutation.",
    ),
    "setPageLayout": (
        "At least one layout field is explicit and a margin object always contains all four sides.",
        "All selected sections are preflighted and observed in normalized points.",
    ),
    "insertBreak": (
        "The handler duplicates and collapses the Body Anchor before insertion.",
        "Only page and explicit section break types are supported and read back.",
    ),
    "save": (
        "Uses only the Binding's existing .docx locator and accepts no path.",
        "An already-saved document may succeed only after the same artifact verification.",
    ),
    "saveAs": (
        "Output is fixed .docx and overwrite policy is explicit with no default.",
        "The same live document remains bound while destination Lease migration has no gap.",
    ),
    "exportPdf": (
        "Exports the complete document to PDF without saving or retargeting the Word document.",
        "Success proves one readable artifact while revision and document state remain unchanged.",
    ),
}


def _content_ranges(value):
    if isinstance(value, dict) or hasattr(value, "items"):
        if {"start", "end", "revision"}.issubset(value):
            yield value
        for item in value.values():
            yield from _content_ranges(item)
    elif isinstance(value, (list, tuple)):
        for item in value:
            yield from _content_ranges(item)


def _range_inside(inner, outer):
    return outer["start"] <= inner["start"] <= inner["end"] <= outer["end"]


def _ranges_are_ordered(ranges):
    return all(
        left["end"] <= right["start"]
        for left, right in zip(ranges, ranges[1:])
    )


def _revision_change_error(result, *, must_change=False, changed_count=None):
    before = result["revisionBefore"]
    after = result["revisionAfter"]
    if must_change and before == after:
        return "revisionAfter must differ from revisionBefore"
    if changed_count is not None:
        if changed_count == 0 and before != after:
            return "a definite no-op must preserve Content Revision"
        if changed_count > 0 and before == after:
            return "an observed change must advance Content Revision"
    return None


def _result_ranges_match_revision(result, revision):
    return all(
        content_range["revision"] == revision
        for content_range in _content_ranges(result)
    )


def _input_revisions(value):
    if isinstance(value, dict) or hasattr(value, "items"):
        revision = value.get("revision")
        if isinstance(revision, str):
            yield revision
        for item in value.values():
            yield from _input_revisions(item)
    elif isinstance(value, (list, tuple)):
        for item in value:
            yield from _input_revisions(item)


def _selected_section_indexes(selector, selected_count):
    if selector["kind"] == "indexes":
        return list(selector["indexes"])
    return list(range(selected_count))


def _anchor_position(anchor):
    if anchor["kind"] == "documentStart":
        return 0
    if anchor["kind"] == "before":
        return anchor["range"]["start"]
    if anchor["kind"] == "after":
        return anchor["range"]["end"]
    return None


def _utf16_units(value):
    return len(value.encode("utf-16-le")) // 2


def _length_in_points(length):
    factors = {
        "pt": 1,
        "in": 72,
        "cm": 72 / 2.54,
        "mm": 72 / 25.4,
    }
    return length["value"] * factors[length["unit"]]


def _observed_length_matches(observed, requested, tolerance=0.5):
    return abs(observed["value"] - _length_in_points(requested)) <= tolerance


def _text_format_patches(action, params):
    if action == "writeContent":
        blocks = params["blocks"]
    elif action == "replaceContent":
        replacement = params["replacement"]
        if replacement["kind"] == "delete":
            return
        if replacement["kind"] == "text":
            blocks = ({"runs": replacement["runs"]},)
        else:
            blocks = replacement["blocks"]
    else:
        return
    for block in blocks:
        for run in block["runs"]:
            if "format" in run:
                yield run["format"]


def _semantic_params_error(action, params):
    if any(
        "fontFamily" in patch
        and (
            "westernFontFamily" in patch
            or "eastAsiaFontFamily" in patch
        )
        for patch in _text_format_patches(action, params)
    ):
        return (
            "fontFamily cannot be combined with script-specific font families"
        )
    if action == "setHeaderFooter" and any(
        update["operation"]["kind"] == "linkToPrevious"
        for update in params["updates"]
    ):
        selector = params["sections"]
        if selector["kind"] == "all" or 0 in selector["indexes"]:
            return "linkToPrevious cannot target section 0"
    return None


def _semantic_result_error(action, params, result):
    expected_input_revision = None
    if action in {"inspectDocument", "findContent"}:
        expected_input_revision = result["revision"]
    elif "revisionBefore" in result:
        expected_input_revision = result["revisionBefore"]
    if expected_input_revision is not None and any(
        revision != expected_input_revision
        for revision in _input_revisions(params)
    ):
        return "every revision-bound input must equal the observed base revision"

    if action == "openDocument":
        if result["artifact"]["path"] != params["path"]:
            return "artifact path must equal the requested document locator"

    elif action == "writeContent":
        error = _revision_change_error(result, must_change=True)
        if error is not None:
            return error
        if result["range"]["revision"] != result["revisionAfter"]:
            return "inserted range must use revisionAfter"
        anchor_position = _anchor_position(params["anchor"])
        if (
            anchor_position is not None
            and result["range"]["start"] != anchor_position
        ):
            return "inserted range must begin at the resolved Body Anchor"

    elif action == "inspectDocument":
        revision = result["revision"]
        if not _result_ranges_match_revision(result, revision):
            return "all inspection ranges must use the snapshot revision"
        if (
            params["scope"]["kind"] == "range"
            and result["scopeRange"] != params["scope"]["range"]
        ):
            return "scopeRange must equal the requested range scope"
        limits = params["limits"]
        if len(result["text"]) > limits["maxTextCharacters"]:
            return "inspection text exceeds maxTextCharacters"
        if len(result["paragraphs"]) > limits["maxParagraphs"]:
            return "inspection paragraphs exceed maxParagraphs"
        run_count = sum(
            len(paragraph["runs"])
            for paragraph in result["paragraphs"]
        )
        if run_count > limits["maxRuns"]:
            return "inspection runs exceed maxRuns"
        structure = result["structure"]
        if structure["sectionCount"] != len(structure["sections"]):
            return "sectionCount must equal the section snapshot count"
        if [section["index"] for section in structure["sections"]] != list(
            range(len(structure["sections"]))
        ):
            return "section snapshots must use complete ordered indexes"
        expected_stories = [
            (area, variant)
            for area in ("header", "footer")
            for variant in ("primary", "firstPage", "evenPages")
        ]
        for section in structure["sections"]:
            observed = [
                (story["area"], story["variant"])
                for story in section["headerFooter"]["stories"]
            ]
            if observed != expected_stories:
                return "section stories must use the canonical six-item order"
        scope_range = result["scopeRange"]
        returned_range = result["returnedRange"]
        if not _range_inside(returned_range, scope_range):
            return "returnedRange must be inside scopeRange"
        remaining = result["remainingRange"]
        if result["truncated"] != (remaining is not None):
            return "truncated must exactly reflect remainingRange"
        if remaining is None:
            if returned_range != scope_range:
                return "an untruncated returnedRange must equal scopeRange"
        elif (
            remaining["start"] != returned_range["end"]
            or remaining["end"] != scope_range["end"]
            or remaining["start"] >= remaining["end"]
            or not _range_inside(remaining, scope_range)
        ):
            return "remainingRange must be the contiguous scope suffix"
        if any(
            not _range_inside(paragraph["range"], returned_range)
            for paragraph in result["paragraphs"]
        ):
            return "paragraph ranges must be inside returnedRange"
        paragraph_ranges = [
            paragraph["range"] for paragraph in result["paragraphs"]
        ]
        if not _ranges_are_ordered(paragraph_ranges):
            return "paragraph ranges must be ordered and non-overlapping"
        for paragraph in result["paragraphs"]:
            run_ranges = [run["range"] for run in paragraph["runs"]]
            if not _ranges_are_ordered(run_ranges) or any(
                not _range_inside(run_range, paragraph["range"])
                for run_range in run_ranges
            ):
                return "run ranges must be ordered inside their paragraph"

    elif action == "findContent":
        revision = result["revision"]
        if not _result_ranges_match_revision(result, revision):
            return "all find ranges must use the result revision"
        requested_scope = params["query"]["scope"]
        if (
            requested_scope["kind"] == "range"
            and result["scopeRange"] != requested_scope["range"]
        ):
            return "scopeRange must equal the requested range scope"
        if len(result["matches"]) > params["limit"]:
            return "returned matches exceed the requested limit"
        ranges = [match["range"] for match in result["matches"]]
        if not _ranges_are_ordered(ranges):
            return "matches must be ordered and non-overlapping"
        if any(
            not _range_inside(match_range, result["scopeRange"])
            for match_range in ranges
        ):
            return "match ranges must be inside scopeRange"
        query_text = params["query"]["text"]
        for match in result["matches"]:
            observed_text = match["text"]
            if match["range"]["end"] - match["range"]["start"] != (
                _utf16_units(observed_text)
            ):
                return "each match range must span its observed UTF-16 text"
            if params["query"]["caseSensitive"]:
                text_matches = observed_text == query_text
            else:
                text_matches = observed_text.casefold() == query_text.casefold()
            if not text_matches:
                return "each observed match must equal the literal query"
        remaining = result["remainingRange"]
        if result["truncated"] != (remaining is not None):
            return "truncated must exactly reflect remainingRange"
        if remaining is not None:
            if len(ranges) != params["limit"] or not ranges:
                return "a truncated find must fill the requested match limit"
            if (
                not _range_inside(remaining, result["scopeRange"])
                or remaining["start"] != ranges[-1]["end"]
                or remaining["end"] != result["scopeRange"]["end"]
                or remaining["start"] >= remaining["end"]
            ):
                return "remainingRange must be the contiguous scope suffix"

    elif action == "replaceContent":
        matched = result["matchedCount"]
        changed = result["changedCount"]
        expected = (
            1
            if params["target"]["kind"] == "range"
            else params["target"]["expectedMatchCount"]
        )
        if matched != expected:
            return "matchedCount must equal the target precondition"
        if changed > matched:
            return "changedCount must not exceed matchedCount"
        if len(result["ranges"]) != matched:
            return "ranges length must equal matchedCount"
        if (
            params["target"]["kind"] == "range"
            and result["ranges"][0]["start"]
            != params["target"]["range"]["start"]
        ):
            return "an exact-range replacement must preserve its start position"
        if params["replacement"]["kind"] == "delete":
            if changed != matched:
                return "deleting non-empty matched targets must change every target"
            if any(
                content_range["start"] != content_range["end"]
                for content_range in result["ranges"]
            ):
                return "deleted targets must return collapsed final ranges"
        elif any(
            content_range["start"] >= content_range["end"]
            for content_range in result["ranges"]
        ):
            return "non-empty replacements must return non-empty final ranges"
        if (
            params["target"]["kind"] == "query"
            and params["replacement"]["kind"] == "text"
        ):
            query = params["target"]["query"]
            replacement_text = "".join(
                run["text"] for run in params["replacement"]["runs"]
            )
            definitely_different = (
                replacement_text != query["text"]
                if query["caseSensitive"]
                else replacement_text.casefold() != query["text"].casefold()
            )
            if definitely_different and changed != matched:
                return "a different literal replacement must change every match"
            replacement_length = _utf16_units(replacement_text)
            if any(
                content_range["end"] - content_range["start"]
                != replacement_length
                for content_range in result["ranges"]
            ):
                return "query replacement ranges must span the replacement text"
        error = _revision_change_error(result, changed_count=changed)
        if error is not None:
            return error
        if not _result_ranges_match_revision(result, result["revisionAfter"]):
            return "replacement ranges must use revisionAfter"
        if not _ranges_are_ordered(result["ranges"]):
            return "replacement ranges must be ordered and non-overlapping"

    elif action == "insertTable":
        error = _revision_change_error(result, must_change=True)
        if error is not None:
            return error
        table = result["table"]
        if table["rowCount"] != len(table["data"]):
            return "rowCount must equal the observed data row count"
        if table["columnCount"] != len(table["data"][0]):
            return "columnCount must equal the observed data width"
        observed_data = tuple(tuple(row) for row in table["data"])
        requested_data = tuple(tuple(row) for row in params["data"])
        if observed_data != requested_data:
            return "observed table data must equal the requested matrix"
        if table["headerRow"] != params["headerRow"]:
            return "observed table header state must equal headerRow"
        if table["range"]["revision"] != result["revisionAfter"]:
            return "table range must use revisionAfter"
        anchor_position = _anchor_position(params["anchor"])
        if (
            anchor_position is not None
            and table["range"]["start"] != anchor_position
        ):
            return "table range must begin at the resolved Body Anchor"

    elif action == "insertImage":
        error = _revision_change_error(result, must_change=True)
        if error is not None:
            return error
        image_range = result["image"].get(
            "range",
            result["image"].get("anchorRange"),
        )
        if image_range["revision"] != result["revisionAfter"]:
            return "image range must use revisionAfter"
        image = result["image"]
        placement = params["placement"]
        if image["kind"] != placement["kind"]:
            return "observed image placement kind must equal the request"
        if image["alternativeText"] != params["alternativeText"]:
            return "observed alternative text must equal the request"
        anchor_position = _anchor_position(params["anchor"])
        if anchor_position is not None and image_range["start"] != anchor_position:
            return "image range must begin at the resolved Body Anchor"
        if placement["kind"] == "floating":
            if image["wrap"] != placement["wrap"]:
                return "observed image wrap must equal the request"
            for axis in ("horizontal", "vertical"):
                if image[axis]["relativeTo"] != placement[axis]["relativeTo"]:
                    return f"observed image {axis} reference must equal the request"
                if not _observed_length_matches(
                    image[axis]["offset"],
                    placement[axis]["offset"],
                ):
                    return f"observed image {axis} offset must equal the request"
        size_request = params["size"]
        observed_size = image["size"]
        if size_request["kind"] == "width" and not _observed_length_matches(
            observed_size["width"],
            size_request["width"],
        ):
            return "observed image width must equal the request"
        if size_request["kind"] == "height" and not _observed_length_matches(
            observed_size["height"],
            size_request["height"],
        ):
            return "observed image height must equal the request"
        if size_request["kind"] == "box":
            width_limit = _length_in_points(size_request["width"])
            height_limit = _length_in_points(size_request["height"])
            if size_request["fit"] == "stretch":
                if (
                    abs(observed_size["width"]["value"] - width_limit) > 0.5
                    or abs(observed_size["height"]["value"] - height_limit)
                    > 0.5
                ):
                    return "stretched image dimensions must equal the requested box"
            elif (
                observed_size["width"]["value"] > width_limit + 0.5
                or observed_size["height"]["value"] > height_limit + 0.5
            ):
                return "contained image dimensions must fit the requested box"

    elif action in {"setHeaderFooter", "setPageLayout"}:
        changed = result["changedCount"]
        error = _revision_change_error(result, changed_count=changed)
        if error is not None:
            return error
        selected_count = result["selectedSectionCount"]
        if changed > selected_count:
            return "changedCount must not exceed selectedSectionCount"
        selected = params["sections"]
        selected_indexes = _selected_section_indexes(
            selected,
            selected_count,
        )
        if selected["kind"] == "indexes" and selected_count != len(
            selected["indexes"]
        ):
            return "selectedSectionCount must equal the requested index count"
        if action == "setHeaderFooter":
            area_order = {"header": 0, "footer": 1}
            variant_order = {
                "primary": 0,
                "firstPage": 1,
                "evenPages": 2,
            }
            update_pairs = sorted(
                (
                    (update["area"], update["variant"])
                    for update in params["updates"]
                ),
                key=lambda pair: (
                    area_order[pair[0]],
                    variant_order[pair[1]],
                ),
            )
            expected_stories = [
                (section_index, area, variant)
                for section_index in selected_indexes
                for area, variant in update_pairs
            ]
            observed_stories = [
                (
                    story["sectionIndex"],
                    story["area"],
                    story["variant"],
                )
                for story in result["stories"]
            ]
            if observed_stories != expected_stories:
                return (
                    "stories must contain every selected section/update pair "
                    "once in canonical order"
                )
            operations = {
                (update["area"], update["variant"]): update["operation"]
                for update in params["updates"]
            }
            for story in result["stories"]:
                operation = operations[(story["area"], story["variant"])]
                if story["variant"] != "primary" and not story[
                    "variantEnabled"
                ]:
                    return "first-page and even-page updates must enable their variant"
                if operation["kind"] == "replace" and (
                    story["text"] != operation["text"]
                    or story["linkToPrevious"]
                    or not story["exists"]
                ):
                    return "a replaced story must be unlinked and equal the requested text"
                if operation["kind"] == "clear" and (
                    story["text"] != "" or story["linkToPrevious"]
                ):
                    return "a cleared story must be empty and unlinked"
                if operation["kind"] == "linkToPrevious" and (
                    story["sectionIndex"] == 0
                    or not story["linkToPrevious"]
                ):
                    return "linkToPrevious must link only a non-first section"
        else:
            if selected_count != len(result["sections"]):
                return "selectedSectionCount must equal the section result count"
            observed_indexes = [
                section["index"] for section in result["sections"]
            ]
            if observed_indexes != selected_indexes:
                return "result sections must equal the complete selection"
            requested_layout = params["layout"]
            for section in result["sections"]:
                observed_layout = section["layout"]
                if (
                    "orientation" in requested_layout
                    and observed_layout["orientation"]
                    != requested_layout["orientation"]
                ):
                    return "observed orientation must equal the request"
                if "margins" in requested_layout:
                    for side in ("top", "right", "bottom", "left"):
                        if not _observed_length_matches(
                            observed_layout["margins"][side],
                            requested_layout["margins"][side],
                        ):
                            return "observed margins must equal the request"

    elif action == "insertBreak":
        error = _revision_change_error(result, must_change=True)
        if error is not None:
            return error
        observed = result["break"]
        if observed["type"] != params["type"]:
            return "observed break type must equal the requested type"
        if observed["range"]["revision"] != result["revisionAfter"]:
            return "break range must use revisionAfter"
        anchor_position = _anchor_position(params["anchor"])
        if (
            anchor_position is not None
            and observed["range"]["start"] != anchor_position
        ):
            return "break range must begin at the resolved Body Anchor"
        if observed["type"] == "page":
            if observed["sectionCountBefore"] != observed["sectionCountAfter"]:
                return "page break must preserve section count"
        elif observed["sectionCountAfter"] != observed["sectionCountBefore"] + 1:
            return "section break must add exactly one section"
        elif observed["followingSectionIndex"] >= observed["sectionCountAfter"]:
            return "followingSectionIndex must name a resulting section"

    elif action in {"save", "saveAs", "exportPdf"}:
        if result["revisionBefore"] != result["revisionAfter"]:
            return "persistence must preserve the stable Content Revision"
        if action == "save":
            return None
        if result["artifact"]["path"] != params["outputPath"]:
            return "artifact path must equal the authorized output locator"
        if (
            action == "exportPdf"
            and result["documentStateBefore"] != result["documentStateAfter"]
        ):
            return "PDF export must preserve the locator-free document state"
        if (
            params["overwritePolicy"] == "failIfExists"
            and result["replacedExisting"]
        ):
            return "failIfExists cannot report replacedExisting"

    return None


def _contract(
    *,
    name,
    category,
    purpose,
    binding_role,
    risk,
    parameters,
    result,
    stable_errors,
    verification,
):
    role_errors = (
        ("SESSION_DOCUMENT_ALREADY_BOUND",)
        if binding_role == "establish"
        else ("SESSION_DOCUMENT_NOT_BOUND",)
        if binding_role == "required"
        else ()
    )
    errors = tuple(dict.fromkeys(
        COMMON_CONTRACT_ERRORS + role_errors + tuple(stable_errors)
    ))
    prerequisites = (
        "Document Intent is resolved for this Word Action.",
        "The Action Session is UNBOUND.",
    ) if binding_role == "establish" else (
        "The Action Session has one exact live Word Session Document Binding.",
        "The bound document and controller remain provably usable.",
    )
    return ActionContract(
        name=name,
        category=category,
        purpose=purpose,
        binding_role=binding_role,
        risk=risk,
        parameters=parameters,
        result=result,
        stable_errors=errors,
        verification=verification,
        prerequisites=prerequisites,
        constraints=_WORD_CONSTRAINTS[name],
        examples=({"params": _WORD_EXAMPLE_PARAMS[name]},),
        parameter_validator=(
            lambda params, action=name: _semantic_params_error(action, params)
        ),
        semantic_validator=(
            lambda params, result, action=name: _semantic_result_error(
                action,
                params,
                result,
            )
        ),
    )


WORD_TARGET_CONTRACT_SET = ApplicationContractSet(
    application="word",
    format_validators=WORD_FORMAT_VALIDATORS,
    contracts=(
        _contract(
            name="createDocument",
            category="documentEstablishment",
            purpose=(
                "Create one blank live Word document and establish the "
                "Session Document Binding."
            ),
            binding_role="establish",
            risk="write",
            parameters=_object({}, ()),
            result=_object(
                {"revision": CONTENT_REVISION, "documentState": CREATED_DOCUMENT_STATE},
                ("revision", "documentState"),
            ),
            stable_errors=(
                "DOCUMENT_CREATE_FAILED",
                "DOCUMENT_BINDING_UNAVAILABLE",
                "RESPONSE_LOST",
            ),
            verification=(
                "Prove the exact object returned by create is live, unsaved, "
                "Lease-bound, and committed before success becomes observable."
            ),
        ),
        _contract(
            name="openDocument",
            category="documentEstablishment",
            purpose=(
                "Open or reuse one exact user-identified existing .docx and "
                "establish the Session Document Binding."
            ),
            binding_role="establish",
            risk="write",
            parameters=_object({"path": DOCX_PATH}, ("path",)),
            result=_object(
                {
                    "revision": CONTENT_REVISION,
                    "artifact": DOCX_ARTIFACT,
                    "documentState": OPENED_DOCUMENT_STATE,
                },
                ("revision", "artifact", "documentState"),
            ),
            stable_errors=(
                "DOCUMENT_NOT_FOUND",
                "DOCUMENT_ACCESS_DENIED",
                "DOCUMENT_PASSWORD_REQUIRED",
                "DOCUMENT_FORMAT_UNSUPPORTED",
                "DOCUMENT_OPEN_FAILED",
                "DOCUMENT_LEASE_CONFLICT",
                "DOCUMENT_QUARANTINED",
                "DOCUMENT_BINDING_UNAVAILABLE",
                "RESPONSE_LOST",
            ),
            verification=(
                "Prove stable file identity, exact open-or-reused object, "
                "observed .docx artifact/state, and atomic Binding plus Lease."
            ),
        ),
        _contract(
            name="writeContent",
            category="bodyContent",
            purpose="Insert structured body content at an explicit body anchor.",
            binding_role="required",
            risk="write",
            parameters=_object(
                {"anchor": BODY_ANCHOR, "blocks": STRUCTURED_BLOCKS},
                ("anchor", "blocks"),
            ),
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "range": NONEMPTY_CONTENT_RANGE,
                },
                ("revisionBefore", "revisionAfter", "range"),
            ),
            stable_errors=CONTENT_MUTATION_ERRORS + (
                "CONTENT_ANCHOR_NOT_PARAGRAPH_BOUNDARY",
                "CONTENT_FORMAT_UNSUPPORTED",
                "CONTENT_WRITE_FAILED",
                "CONTENT_VERIFICATION_FAILED",
            ),
            verification=(
                "Read back inserted text, block structure, semantic headings, "
                "and every explicitly supplied format field at the fresh range."
            ),
        ),
        _contract(
            name="inspectDocument",
            category="bodyContent",
            purpose=(
                "Return one revision-coherent normalized snapshot of bound "
                "Word body content, structure, layout, and document state."
            ),
            binding_role="required",
            risk="read",
            parameters=_object(
                {
                    "scope": BODY_SCOPE,
                    "limits": _object(
                        {
                            "maxTextCharacters": _integer(
                                minimum=1, maximum=65536
                            ),
                            "maxParagraphs": _integer(
                                minimum=1, maximum=512
                            ),
                            "maxRuns": _integer(
                                minimum=1, maximum=2048
                            ),
                        },
                        (
                            "maxTextCharacters",
                            "maxParagraphs",
                            "maxRuns",
                        ),
                    ),
                },
                ("scope", "limits"),
            ),
            result=_object(
                {
                    "revision": CONTENT_REVISION,
                    "scopeRange": CONTENT_RANGE,
                    "returnedRange": CONTENT_RANGE,
                    "text": _string(
                        maximum=65536,
                        format="wordNormalizedText",
                    ),
                    "paragraphs": _array(
                        PARAGRAPH_SNAPSHOT,
                        maximum=512,
                    ),
                    "structure": STRUCTURE_SNAPSHOT,
                    "documentState": DOCUMENT_STATE,
                    "truncated": {"type": "boolean"},
                    "remainingRange": _nullable(CONTENT_RANGE),
                },
                (
                    "revision",
                    "scopeRange",
                    "returnedRange",
                    "text",
                    "paragraphs",
                    "structure",
                    "documentState",
                    "truncated",
                    "remainingRange",
                ),
            ),
            stable_errors=CONTENT_READ_ERRORS + (
                "CONTENT_READ_FAILED",
                "STALE_DOCUMENT_REVISION",
            ),
            verification=(
                "Take all body, section, layout, header/footer, and persistence "
                "facts from one coherent revision; never return mixed snapshots."
            ),
        ),
        _contract(
            name="findContent",
            category="bodyContent",
            purpose=(
                "Find bounded, literal, non-overlapping body-text matches "
                "against one coherent content revision."
            ),
            binding_role="required",
            risk="read",
            parameters=_object(
                {
                    "query": TEXT_QUERY,
                    "limit": _integer(minimum=1, maximum=200),
                },
                ("query", "limit"),
            ),
            result=_object(
                {
                    "revision": CONTENT_REVISION,
                    "scopeRange": CONTENT_RANGE,
                    "matches": _array(
                        _object(
                            {
                                "range": NONEMPTY_CONTENT_RANGE,
                                "text": _string(
                                    minimum=1,
                                    maximum=4096,
                                    format="wordQueryText",
                                ),
                            },
                            ("range", "text"),
                        ),
                        maximum=200,
                    ),
                    "truncated": {"type": "boolean"},
                    "remainingRange": _nullable(CONTENT_RANGE),
                },
                (
                    "revision",
                    "scopeRange",
                    "matches",
                    "truncated",
                    "remainingRange",
                ),
            ),
            stable_errors=CONTENT_READ_ERRORS + ("CONTENT_READ_FAILED",),
            verification=(
                "Return document-order non-overlapping observed literals and "
                "ranges from one revision; an empty match list is success."
            ),
        ),
        _contract(
            name="replaceContent",
            category="bodyContent",
            purpose=(
                "Replace one exact range or an explicitly counted literal "
                "query result with preflighted structured content."
            ),
            binding_role="required",
            risk="destructive",
            parameters={
                "oneOf": (
                    _object(
                        {
                            "target": _object(
                                {
                                    "kind": _const("range"),
                                    "range": NONEMPTY_CONTENT_RANGE,
                                },
                                ("kind", "range"),
                            ),
                            "replacement": {
                                "oneOf": (
                                    _object(
                                        {"kind": _const("delete")},
                                        ("kind",),
                                    ),
                                    _object(
                                        {
                                            "kind": _const("blocks"),
                                            "blocks": STRUCTURED_BLOCKS,
                                        },
                                        ("kind", "blocks"),
                                    ),
                                )
                            },
                        },
                        ("target", "replacement"),
                    ),
                    _object(
                        {
                            "target": _object(
                                {
                                    "kind": _const("query"),
                                    "query": TEXT_QUERY,
                                    "expectedMatchCount": _integer(
                                        minimum=1, maximum=200
                                    ),
                                },
                                (
                                    "kind",
                                    "query",
                                    "expectedMatchCount",
                                ),
                            ),
                            "replacement": {
                                "oneOf": (
                                    _object(
                                        {"kind": _const("delete")},
                                        ("kind",),
                                    ),
                                    _object(
                                        {
                                            "kind": _const("text"),
                                            "runs": TEXT_RUNS,
                                        },
                                        ("kind", "runs"),
                                    ),
                                )
                            },
                        },
                        ("target", "replacement"),
                    ),
                )
            },
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "matchedCount": _integer(minimum=1, maximum=200),
                    "changedCount": _integer(minimum=0, maximum=200),
                    "ranges": _array(
                        CONTENT_RANGE,
                        minimum=1,
                        maximum=200,
                    ),
                },
                (
                    "revisionBefore",
                    "revisionAfter",
                    "matchedCount",
                    "changedCount",
                    "ranges",
                ),
            ),
            stable_errors=CONTENT_MUTATION_ERRORS + (
                "MATCH_COUNT_MISMATCH",
                "CONTENT_RANGE_NOT_REPLACEABLE",
                "CONTENT_FORMAT_UNSUPPORTED",
                "CONTENT_WRITE_FAILED",
                "CONTENT_VERIFICATION_FAILED",
            ),
            verification=(
                "Preflight all targets and counts before the first write, "
                "mutate from the end, then read back every final fresh range."
            ),
        ),
        _contract(
            name="insertTable",
            category="embeddedContent",
            purpose=(
                "Insert one non-empty rectangular plain-text table at an "
                "explicit body anchor."
            ),
            binding_role="required",
            risk="write",
            parameters=_object(
                {
                    "anchor": BODY_ANCHOR,
                    "data": TABLE_DATA,
                    "headerRow": {"type": "boolean"},
                },
                ("anchor", "data", "headerRow"),
            ),
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "table": _object(
                        {
                            "range": NONEMPTY_CONTENT_RANGE,
                            "rowCount": _integer(minimum=1, maximum=100),
                            "columnCount": _integer(
                                minimum=1, maximum=50
                            ),
                            "headerRow": {"type": "boolean"},
                            "data": TABLE_DATA,
                        },
                        (
                            "range",
                            "rowCount",
                            "columnCount",
                            "headerRow",
                            "data",
                        ),
                    ),
                },
                ("revisionBefore", "revisionAfter", "table"),
            ),
            stable_errors=CONTENT_MUTATION_ERRORS + (
                "TABLE_ANCHOR_UNSUPPORTED",
                "TABLE_APPLY_FAILED",
                "TABLE_VERIFICATION_FAILED",
                "WORD_CAPABILITY_UNAVAILABLE",
            ),
            verification=(
                "Read back exact dimensions, repeating-header state, range, "
                "and normalized text from every created cell."
            ),
        ),
        _contract(
            name="insertImage",
            category="embeddedContent",
            purpose=(
                "Validate and embed one PNG or JPEG at an explicit body anchor "
                "with explicit placement, size, and alternative text."
            ),
            binding_role="required",
            risk="write",
            parameters=_object(
                {
                    "anchor": BODY_ANCHOR,
                    "source": _object(
                        {
                            "kind": _const("file"),
                            "path": IMAGE_PATH,
                        },
                        ("kind", "path"),
                    ),
                    "placement": IMAGE_PLACEMENT,
                    "size": IMAGE_SIZE,
                    "alternativeText": ALTERNATIVE_TEXT,
                },
                (
                    "anchor",
                    "source",
                    "placement",
                    "size",
                    "alternativeText",
                ),
            ),
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "image": OBSERVED_IMAGE,
                },
                ("revisionBefore", "revisionAfter", "image"),
            ),
            stable_errors=CONTENT_MUTATION_ERRORS + (
                "IMAGE_SOURCE_NOT_FOUND",
                "IMAGE_SOURCE_UNREADABLE",
                "IMAGE_FORMAT_UNSUPPORTED",
                "IMAGE_SOURCE_LIMIT_EXCEEDED",
                "IMAGE_PLACEMENT_UNSUPPORTED",
                "IMAGE_APPLY_FAILED",
                "IMAGE_VERIFICATION_FAILED",
                "WORD_CAPABILITY_UNAVAILABLE",
            ),
            verification=(
                "Read back embedded media hash/type, dimensions, exact anchor, "
                "placement, wrap, and alternative text before success."
            ),
        ),
        _contract(
            name="setHeaderFooter",
            category="pageStructure",
            purpose=(
                "Apply explicit header/footer operations to revision-bound "
                "sections and read back every touched story."
            ),
            binding_role="required",
            risk="write",
            parameters=_object(
                {
                    "sections": SECTION_SELECTOR,
                    "updates": _array(
                        HEADER_FOOTER_UPDATE,
                        minimum=1,
                        maximum=6,
                        **{"x-maxUtf16Text": 1048576},
                    ),
                },
                ("sections", "updates"),
                **{"x-uniqueBy": ("updates", "area", "variant")},
            ),
            result=_object(
                {
                    "selectedSectionCount": _integer(
                        minimum=1, maximum=64
                    ),
                    "changedCount": _integer(minimum=0, maximum=64),
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "stories": _array(
                        OBSERVED_HEADER_FOOTER_STORY,
                        minimum=1,
                        maximum=384,
                        **{"x-maxUtf16Text": 1048576},
                    ),
                },
                (
                    "selectedSectionCount",
                    "changedCount",
                    "revisionBefore",
                    "revisionAfter",
                    "stories",
                ),
            ),
            stable_errors=MUTATION_ERRORS + (
                "STALE_DOCUMENT_REVISION",
                "SECTION_NOT_FOUND",
                "HEADER_FOOTER_LINK_INVALID",
                "HEADER_FOOTER_APPLY_FAILED",
                "HEADER_FOOTER_VERIFICATION_FAILED",
                "CONTENT_LIMIT_EXCEEDED",
                "CONTENT_CHANGED_DURING_ACTION",
                "WORD_CAPABILITY_UNAVAILABLE",
            ),
            verification=(
                "Preflight every section/update, then read back enabled flags, "
                "link state, existence, and normalized text in stable order."
            ),
        ),
        _contract(
            name="setPageLayout",
            category="pageStructure",
            purpose=(
                "Apply explicit orientation and complete margin changes to "
                "revision-bound sections."
            ),
            binding_role="required",
            risk="write",
            parameters=_object(
                {
                    "sections": SECTION_SELECTOR,
                    "layout": _object(
                        {
                            "orientation": _string(
                                enum=("portrait", "landscape")
                            ),
                            "margins": MARGINS_INPUT,
                        },
                        (),
                        **{"x-atLeastOne": ("orientation", "margins")},
                    ),
                },
                ("sections", "layout"),
            ),
            result=_object(
                {
                    "selectedSectionCount": _integer(
                        minimum=1, maximum=64
                    ),
                    "changedCount": _integer(minimum=0, maximum=64),
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "sections": _array(
                        _object(
                            {
                                "index": _integer(minimum=0),
                                "layout": LAYOUT_SNAPSHOT,
                            },
                            ("index", "layout"),
                        ),
                        minimum=1,
                        maximum=64,
                    ),
                },
                (
                    "selectedSectionCount",
                    "changedCount",
                    "revisionBefore",
                    "revisionAfter",
                    "sections",
                ),
            ),
            stable_errors=MUTATION_ERRORS + (
                "STALE_DOCUMENT_REVISION",
                "SECTION_NOT_FOUND",
                "PAGE_LAYOUT_INVALID",
                "PAGE_LAYOUT_APPLY_FAILED",
                "PAGE_LAYOUT_VERIFICATION_FAILED",
                "CONTENT_CHANGED_DURING_ACTION",
                "CONTENT_LIMIT_EXCEEDED",
                "WORD_CAPABILITY_UNAVAILABLE",
            ),
            verification=(
                "Preflight all section dimensions, then read back normalized "
                "orientation and all four margins for every selected section."
            ),
        ),
        _contract(
            name="insertBreak",
            category="pageStructure",
            purpose=(
                "Insert one explicit page or section break at a collapsed "
                "body anchor and verify the resulting structure."
            ),
            binding_role="required",
            risk="write",
            parameters=_object(
                {
                    "anchor": BODY_ANCHOR,
                    "type": _string(
                        enum=(
                            "page",
                            "sectionNextPage",
                            "sectionContinuous",
                            "sectionEvenPage",
                            "sectionOddPage",
                        )
                    ),
                },
                ("anchor", "type"),
            ),
            result={
                "oneOf": (
                    _object(
                        {
                            "revisionBefore": CONTENT_REVISION,
                            "revisionAfter": CONTENT_REVISION,
                            "break": _object(
                                {
                                    "type": _const("page"),
                                    "range": NONEMPTY_CONTENT_RANGE,
                                    "sectionCountBefore": _integer(
                                        minimum=1
                                    ),
                                    "sectionCountAfter": _integer(
                                        minimum=1
                                    ),
                                },
                                (
                                    "type",
                                    "range",
                                    "sectionCountBefore",
                                    "sectionCountAfter",
                                ),
                            ),
                        },
                        ("revisionBefore", "revisionAfter", "break"),
                    ),
                    _object(
                        {
                            "revisionBefore": CONTENT_REVISION,
                            "revisionAfter": CONTENT_REVISION,
                            "break": _object(
                                {
                                    "type": _string(
                                        enum=(
                                            "sectionNextPage",
                                            "sectionContinuous",
                                            "sectionEvenPage",
                                            "sectionOddPage",
                                        )
                                    ),
                                    "range": NONEMPTY_CONTENT_RANGE,
                                    "sectionCountBefore": _integer(
                                        minimum=1
                                    ),
                                    "sectionCountAfter": _integer(
                                        minimum=1
                                    ),
                                    "followingSectionIndex": _integer(
                                        minimum=1
                                    ),
                                },
                                (
                                    "type",
                                    "range",
                                    "sectionCountBefore",
                                    "sectionCountAfter",
                                    "followingSectionIndex",
                                ),
                            ),
                        },
                        ("revisionBefore", "revisionAfter", "break"),
                    ),
                )
            },
            stable_errors=CONTENT_MUTATION_ERRORS + (
                "BREAK_ANCHOR_UNSUPPORTED",
                "BREAK_APPLY_FAILED",
                "BREAK_VERIFICATION_FAILED",
                "WORD_CAPABILITY_UNAVAILABLE",
            ),
            verification=(
                "Use a duplicated collapsed range, then read back the page "
                "marker or exact section count/start type at a fresh revision."
            ),
        ),
        _contract(
            name="save",
            category="persistence",
            purpose=(
                "Persist the bound .docx through its existing locator and "
                "verify the saved artifact."
            ),
            binding_role="required",
            risk="write",
            parameters=_object({}, ()),
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "artifact": DOCX_ARTIFACT,
                    "documentState": SAVED_DOCUMENT_STATE,
                },
                (
                    "revisionBefore",
                    "revisionAfter",
                    "artifact",
                    "documentState",
                ),
            ),
            stable_errors=MUTATION_ERRORS + (
                "PERSISTENCE_LOCATOR_REQUIRED",
                "OUTPUT_ACCESS_DENIED",
                "OUTPUT_WRITE_FAILED",
                "OUTPUT_VERIFICATION_FAILED",
                "DOCUMENT_CHANGED_DURING_ACTION",
            ),
            verification=(
                "Prove the same backing identity, saved state, .docx format, "
                "ordinary non-empty artifact, and live exact Binding."
            ),
        ),
        _contract(
            name="saveAs",
            category="persistence",
            purpose=(
                "Persist the bound live document as .docx at one explicit "
                "user-authorized output locator."
            ),
            binding_role="required",
            risk="destructive",
            parameters=_object(
                {
                    "outputPath": DOCX_PATH,
                    "overwritePolicy": OVERWRITE_POLICY,
                },
                ("outputPath", "overwritePolicy"),
            ),
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "artifact": DOCX_ARTIFACT,
                    "documentState": SAVED_DOCUMENT_STATE,
                    "replacedExisting": {"type": "boolean"},
                },
                (
                    "revisionBefore",
                    "revisionAfter",
                    "artifact",
                    "documentState",
                    "replacedExisting",
                ),
            ),
            stable_errors=BINDING_ERRORS + (
                "OUTPUT_ALREADY_EXISTS",
                "OUTPUT_MATCHES_BOUND_DOCUMENT",
                "OUTPUT_PARENT_NOT_FOUND",
                "OUTPUT_PATH_INVALID",
                "OUTPUT_IN_USE",
                "OUTPUT_ACCESS_DENIED",
                "OUTPUT_WRITE_FAILED",
                "OUTPUT_VERIFICATION_FAILED",
                "DOCUMENT_LEASE_CONFLICT",
                "DOCUMENT_QUARANTINED",
                "DOCUMENT_CHANGED_DURING_ACTION",
            ),
            verification=(
                "Prove explicit .docx format, artifact, saved state, same live "
                "document, destination identity, and gap-free Lease migration."
            ),
        ),
        _contract(
            name="exportPdf",
            category="persistence",
            purpose=(
                "Export the complete bound Word document to one explicit "
                "user-authorized PDF artifact."
            ),
            binding_role="required",
            risk="destructive",
            parameters=_object(
                {
                    "outputPath": PDF_PATH,
                    "overwritePolicy": OVERWRITE_POLICY,
                },
                ("outputPath", "overwritePolicy"),
            ),
            result=_object(
                {
                    "revisionBefore": CONTENT_REVISION,
                    "revisionAfter": CONTENT_REVISION,
                    "artifact": PDF_ARTIFACT,
                    "documentStateBefore": DOCUMENT_STATE,
                    "documentStateAfter": DOCUMENT_STATE,
                    "replacedExisting": {"type": "boolean"},
                },
                (
                    "revisionBefore",
                    "revisionAfter",
                    "artifact",
                    "documentStateBefore",
                    "documentStateAfter",
                    "replacedExisting",
                ),
            ),
            stable_errors=BINDING_ERRORS + (
                "OUTPUT_ALREADY_EXISTS",
                "OUTPUT_PARENT_NOT_FOUND",
                "OUTPUT_PATH_INVALID",
                "OUTPUT_IN_USE",
                "OUTPUT_ACCESS_DENIED",
                "OUTPUT_WRITE_FAILED",
                "OUTPUT_VERIFICATION_FAILED",
                "DOCUMENT_CHANGED_DURING_ACTION",
            ),
            verification=(
                "Prove a non-empty readable PDF for one stable content revision "
                "without saving, retargeting, or changing the Word Binding."
            ),
        ),
    ),
)

WORD_TARGET_ACTION_INDEX = WORD_TARGET_CONTRACT_SET.action_index()


# Production deliberately omits only saveAs, whose same-object locator and
# Lease migration remains a separate architecture milestone.  Every admitted
# Action below has a real handler and live Writer Backend implementation.
_WORD_PRODUCTION_ACTIONS = frozenset({
    "createDocument",
    "openDocument",
    "writeContent",
    "inspectDocument",
    "findContent",
    "replaceContent",
    "insertTable",
    "insertImage",
    "setHeaderFooter",
    "setPageLayout",
    "insertBreak",
    "save",
    "exportPdf",
})

WORD_PRODUCTION_CONTRACT_SET = ApplicationContractSet(
    application="word",
    contracts=tuple(
        contract
        for contract in WORD_TARGET_CONTRACT_SET.contracts
        if contract.name in _WORD_PRODUCTION_ACTIONS
    ),
    format_validators=WORD_FORMAT_VALIDATORS,
)

WORD_PRODUCTION_ACTION_INDEX = (
    WORD_PRODUCTION_CONTRACT_SET.action_index()
)
