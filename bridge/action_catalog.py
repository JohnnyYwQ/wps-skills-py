"""Machine-readable Windows Action Contracts and standard-library validation."""

import json
import re
from pathlib import Path

from powershell_contracts import CHART_TYPE_ACTIONS


MANIFEST_PATH = Path(__file__).with_name("action_manifest.json")
SUPPORTED_SCHEMA_KEYWORDS = {
    "type",
    "description",
    "properties",
    "required",
    "additionalProperties",
    "enum",
    "items",
    "minimum",
    "maximum",
    "pattern",
    "anyOf",
}
SUPPORTED_SCHEMA_TYPES = {
    "object", "array", "string", "number", "integer", "boolean", "null"
}
SUPPORTED_OWNERS = {"bridge", "excel", "ppt", "word"}
SUPPORTED_RISKS = {"read", "write", "destructive"}
INTERNAL_WINDOWS_HANDLERS = {
    "excel": frozenset({"ping"}),
    "ppt": frozenset({"ping"}),
    "word": frozenset({"ping"}),
}
TOP_LEVEL_FIELDS = {"schema_version", "actions"}
CONTRACT_FIELDS = {
    "owner",
    "action",
    "description",
    "parameters",
    "result",
    "prerequisites",
    "risk",
    "examples",
}
REQUIRED_CONTRACT_FIELDS = CONTRACT_FIELDS - {"examples"}


class ActionManifestError(ValueError):
    code = "INVALID_ACTION_MANIFEST"


class ActionValidationError(ValueError):
    """A runtime parameter or result does not match an Action Contract."""


class UnknownActionError(LookupError):
    code = "UNKNOWN_ACTION"


class AmbiguousActionError(LookupError):
    code = "AMBIGUOUS_ACTION"

    def __init__(self, action, owners):
        self.action = action
        self.owners = list(owners)
        super().__init__(f"action '{action}' has candidate owners {self.owners}")


class ActionNotSupportedError(LookupError):
    code = "ACTION_NOT_SUPPORTED_FOR_APP"

    def __init__(self, owner, action, supported_owners):
        self.owner = owner
        self.action = action
        self.supported_owners = list(supported_owners)
        super().__init__(f"owner '{owner}' does not support action '{action}'")


def _manifest_error(path, message):
    raise ActionManifestError(f"{path} {message}")


def _validate_schema_keywords(schema, path):
    if not isinstance(schema, dict):
        _manifest_error(path, "must be object")
    for keyword in schema:
        if keyword not in SUPPORTED_SCHEMA_KEYWORDS:
            _manifest_error(f"{path}.{keyword}", "is not supported")

    schema_type = schema.get("type")
    if schema_type is not None and (
        not isinstance(schema_type, str)
        or schema_type not in SUPPORTED_SCHEMA_TYPES
    ):
        _manifest_error(f"{path}.type", f"must be one of {sorted(SUPPORTED_SCHEMA_TYPES)}")
    if schema_type is None and "anyOf" not in schema:
        _manifest_error(path, "must declare type or anyOf")
    description = schema.get("description")
    if description is not None and not isinstance(description, str):
        _manifest_error(f"{path}.description", "must be string")

    properties = schema.get("properties", {})
    if not isinstance(properties, dict):
        _manifest_error(f"{path}.properties", "must be object")
    if "properties" in schema and schema_type != "object":
        _manifest_error(f"{path}.properties", "requires type object")
    for name, child in properties.items():
        if not isinstance(name, str) or not name:
            _manifest_error(f"{path}.properties", "keys must be non-empty strings")
        _validate_schema_keywords(child, f"{path}.properties.{name}")

    required = schema.get("required", [])
    if not isinstance(required, list) or any(not isinstance(value, str) for value in required):
        _manifest_error(f"{path}.required", "must be an array of strings")
    if len(set(required)) != len(required):
        _manifest_error(f"{path}.required", "must not contain duplicates")
    unknown_required = set(required) - set(properties)
    if unknown_required:
        _manifest_error(f"{path}.required", f"contains unknown properties {sorted(unknown_required)}")
    if "required" in schema and schema_type != "object":
        _manifest_error(f"{path}.required", "requires type object")

    if "additionalProperties" in schema:
        if not isinstance(schema["additionalProperties"], bool):
            _manifest_error(f"{path}.additionalProperties", "must be boolean")
        if schema_type != "object":
            _manifest_error(f"{path}.additionalProperties", "requires type object")
    if "items" in schema:
        if schema_type != "array":
            _manifest_error(f"{path}.items", "requires type array")
        _validate_schema_keywords(schema["items"], f"{path}.items")
    any_of = schema.get("anyOf", [])
    if "anyOf" in schema and (not isinstance(any_of, list) or not any_of):
        _manifest_error(f"{path}.anyOf", "must be a non-empty array")
    for index, child in enumerate(any_of):
        _validate_schema_keywords(child, f"{path}.anyOf[{index}]")

    enum = schema.get("enum")
    if "enum" in schema and (not isinstance(enum, list) or not enum):
        _manifest_error(f"{path}.enum", "must be a non-empty array")
    if isinstance(enum, list) and schema_type is not None:
        for index, value in enumerate(enum):
            if not _matches_type(value, schema_type):
                _manifest_error(
                    f"{path}.enum[{index}]",
                    f"must be {schema_type}",
                )
    for keyword in ("minimum", "maximum"):
        if keyword in schema:
            value = schema[keyword]
            if isinstance(value, bool) or not isinstance(value, (int, float)):
                _manifest_error(f"{path}.{keyword}", "must be number")
            if schema_type not in {"number", "integer"}:
                _manifest_error(f"{path}.{keyword}", "requires numeric type")
    if "minimum" in schema and "maximum" in schema:
        if schema["minimum"] > schema["maximum"]:
            _manifest_error(path, "minimum must not exceed maximum")
    if "pattern" in schema:
        if schema_type != "string" or not isinstance(schema["pattern"], str):
            _manifest_error(f"{path}.pattern", "requires string type and value")
        try:
            re.compile(schema["pattern"])
        except re.error as exc:
            _manifest_error(f"{path}.pattern", f"is invalid: {exc}")


def _validate_contract(contract, index):
    path = f"actions[{index}]"
    if not isinstance(contract, dict):
        _manifest_error(path, "must be object")
    missing = REQUIRED_CONTRACT_FIELDS - set(contract)
    if missing:
        _manifest_error(path, f"is missing fields {sorted(missing)}")
    unknown = set(contract) - CONTRACT_FIELDS
    if unknown:
        _manifest_error(path, f"contains unknown fields {sorted(unknown)}")
    if not isinstance(contract["owner"], str) or contract["owner"] not in SUPPORTED_OWNERS:
        _manifest_error(f"{path}.owner", f"must be one of {sorted(SUPPORTED_OWNERS)}")
    if not isinstance(contract["action"], str) or not re.fullmatch(r"[A-Za-z][A-Za-z0-9]*", contract["action"]):
        _manifest_error(f"{path}.action", "must be a wire Action name")
    if not isinstance(contract["description"], str) or not contract["description"].strip():
        _manifest_error(f"{path}.description", "must be a non-empty string")
    if not isinstance(contract["risk"], str) or contract["risk"] not in SUPPORTED_RISKS:
        _manifest_error(f"{path}.risk", f"must be one of {sorted(SUPPORTED_RISKS)}")
    prerequisites = contract["prerequisites"]
    if not isinstance(prerequisites, list) or any(
        not isinstance(value, str) or not value for value in prerequisites
    ):
        _manifest_error(f"{path}.prerequisites", "must be an array of non-empty strings")
    _validate_schema_keywords(contract["parameters"], f"{path}.parameters")
    _validate_schema_keywords(contract["result"], f"{path}.result")
    for field in ("parameters", "result"):
        if contract[field].get("type") != "object":
            _manifest_error(f"{path}.{field}", "must have object type")
    examples = contract.get("examples", [])
    if not isinstance(examples, list):
        _manifest_error(f"{path}.examples", "must be array")
    for example_index, example in enumerate(examples):
        example_path = f"{path}.examples[{example_index}]"
        if not isinstance(example, dict) or set(example) != {"params", "result"}:
            _manifest_error(example_path, "must contain exactly params and result")
        for field, schema_name in (("params", "parameters"), ("result", "result")):
            try:
                validate_instance(example[field], contract[schema_name], f"{example_path}.{field}")
            except ActionValidationError as exc:
                _manifest_error(str(exc), "")


def load_action_manifest(path=MANIFEST_PATH):
    path = Path(path)
    try:
        manifest = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ActionManifestError(f"manifest could not be loaded: {exc}") from exc
    if not isinstance(manifest, dict):
        _manifest_error("manifest", "must be object")
    unknown = set(manifest) - TOP_LEVEL_FIELDS
    if unknown:
        _manifest_error("manifest", f"contains unknown fields {sorted(unknown)}")
    if type(manifest.get("schema_version")) is not int or manifest["schema_version"] != 1:
        _manifest_error("schema_version", "must be 1")
    actions = manifest.get("actions")
    if not isinstance(actions, list):
        _manifest_error("actions", "must be array")
    identities = set()
    for index, contract in enumerate(actions):
        _validate_contract(contract, index)
        identity = (contract["owner"], contract["action"])
        if identity in identities:
            _manifest_error(f"actions[{index}]", f"duplicate contract {identity}")
        identities.add(identity)
    return manifest


def _matches_type(value, expected):
    if expected == "object":
        return isinstance(value, dict)
    if expected == "array":
        return isinstance(value, list)
    if expected == "string":
        return isinstance(value, str)
    if expected == "number":
        return isinstance(value, (int, float)) and not isinstance(value, bool)
    if expected == "integer":
        return isinstance(value, int) and not isinstance(value, bool)
    if expected == "boolean":
        return isinstance(value, bool)
    return value is None


def validate_instance(value, schema, path="value"):
    """Validate one JSON-compatible value against the supported schema subset."""
    if "anyOf" in schema:
        for alternative in schema["anyOf"]:
            try:
                validate_instance(value, alternative, path)
                break
            except ActionValidationError:
                continue
        else:
            raise ActionValidationError(f"{path} must match anyOf")

    expected = schema.get("type")
    if expected is not None and not _matches_type(value, expected):
        raise ActionValidationError(f"{path} must be {expected}")
    if "enum" in schema and value not in schema["enum"]:
        raise ActionValidationError(f"{path} must be one of {schema['enum']}")
    if "minimum" in schema and value < schema["minimum"]:
        raise ActionValidationError(f"{path} must be >= {schema['minimum']}")
    if "maximum" in schema and value > schema["maximum"]:
        raise ActionValidationError(f"{path} must be <= {schema['maximum']}")
    if "pattern" in schema and re.search(schema["pattern"], value) is None:
        raise ActionValidationError(f"{path} must match {schema['pattern']}")

    if expected == "object":
        properties = schema.get("properties", {})
        for name in schema.get("required", []):
            if name not in value:
                raise ActionValidationError(f"{path}.{name} is required")
        if schema.get("additionalProperties") is False:
            for name in value:
                if name not in properties:
                    raise ActionValidationError(f"{path}.{name} is not allowed")
        for name, child in properties.items():
            if name in value:
                validate_instance(value[name], child, f"{path}.{name}")
    elif expected == "array" and "items" in schema:
        for index, item in enumerate(value):
            validate_instance(item, schema["items"], f"{path}[{index}]")


class ActionCatalog:
    """Validated, indexed Action Contracts keyed by ``(owner, action)``."""

    def __init__(self, manifest):
        self.schema_version = manifest["schema_version"]
        self._contracts = {
            (contract["owner"], contract["action"]): contract
            for contract in manifest["actions"]
        }

    @classmethod
    def from_path(cls, path=MANIFEST_PATH):
        return cls(load_action_manifest(path))

    def get(self, owner, action):
        try:
            return self._contracts[(owner, action)]
        except KeyError as exc:
            owners = self.owners_for(action)
            if owners:
                raise ActionNotSupportedError(owner, action, owners) from exc
            raise UnknownActionError(f"unknown action '{action}'") from exc

    def owners_for(self, action):
        return sorted(owner for owner, name in self._contracts if name == action)

    def describe(self, action, owner=None):
        if owner is not None:
            return self.get(owner, action)
        owners = self.owners_for(action)
        if not owners:
            raise UnknownActionError(f"unknown action '{action}'")
        if len(owners) > 1:
            raise AmbiguousActionError(action, owners)
        return self._contracts[(owners[0], action)]

    def list(self, owner=None):
        summaries = []
        for (contract_owner, _action), contract in sorted(self._contracts.items()):
            if owner is not None and owner != contract_owner:
                continue
            summaries.append({
                "owner": contract_owner,
                "action": contract["action"],
                "description": contract["description"],
                "risk": contract["risk"],
            })
        return summaries

    def search(self, query, owner=None):
        needle = query.casefold()
        return [
            item for item in self.list(owner=owner)
            if needle in item["action"].casefold()
            or needle in item["description"].casefold()
        ]

    def validate_params(self, owner, action, params):
        contract = self.get(owner, action)
        validate_instance(params, contract["parameters"], "params")

    def validate_result(self, owner, action, result):
        contract = self.get(owner, action)
        validate_instance(result, contract["result"], "result")


def validate_windows_implementation_consistency(catalog, scripts_by_owner):
    """Repository check only: compare contracts with embedded ``Exec-*`` handlers."""
    for owner, script in scripts_by_owner.items():
        if owner not in INTERNAL_WINDOWS_HANDLERS:
            _manifest_error(f"implementation.{owner}", "has no handler exclusion policy")
        implemented = set(re.findall(r"function Exec-(\w+)\b", script))
        public_implemented = implemented - INTERNAL_WINDOWS_HANDLERS[owner]
        contracted = {
            item["action"] for item in catalog.list(owner=owner)
        }
        missing_contracts = sorted(public_implemented - contracted)
        missing_handlers = sorted(contracted - public_implemented)
        if missing_contracts or missing_handlers:
            _manifest_error(
                f"implementation.{owner}",
                f"differs from manifest: missing contracts={missing_contracts}, "
                f"missing handlers={missing_handlers}",
            )

        if owner in CHART_TYPE_ACTIONS:
            contract = catalog.get(owner, CHART_TYPE_ACTIONS[owner])
            chart_schema = contract["parameters"]["properties"]["chartType"]
            string_schema = next(
                option for option in chart_schema["anyOf"]
                if option.get("type") == "string"
            )
            expected_types = set(string_schema["enum"])
            function_match = re.search(
                r"(?ms)^function Convert-ChartType\(\$value\) \{.*?^\}",
                script,
            )
            implemented_types = set()
            if function_match:
                implemented_types = set(re.findall(
                    r'^\s*"([A-Za-z]+)"\s*\{\s*return\b',
                    function_match.group(0),
                    re.MULTILINE,
                ))
            if implemented_types != expected_types:
                _manifest_error(
                    f"implementation.{owner}.chartType",
                    "differs from manifest: "
                    f"manifest={sorted(expected_types)}, "
                    f"implementation={sorted(implemented_types)}",
                )
