"""Shared application-scoped Action Session and document-binding Core."""

from dataclasses import dataclass
import math
from pathlib import Path
import re
import threading
import time
from types import MappingProxyType
from typing import (
    Any,
    Callable,
    Mapping,
    Optional,
    Protocol,
    Sequence,
    runtime_checkable,
)
import uuid


_APPLICATIONS = frozenset({"excel", "ppt", "word"})
_BINDING_ROLES = frozenset({"none", "establish", "required"})
_RISKS = frozenset({"read", "write", "destructive"})
_OUTCOMES = frozenset({"succeeded", "failed", "unknown"})
_CONTROLLER_STATES = frozenset({"usable", "broken"})
_CONTROL_PLANE_ERROR_CODES = frozenset({"RUNTIME_CLOSED"})


class ContractValidationError(ValueError):
    def __init__(self, *, action: str, location: str, message: str):
        super().__init__(f"{action} {location}: {message}")
        self.action = action
        self.location = location
        self.message = message


def _freeze_json(value):
    if isinstance(value, Mapping):
        return MappingProxyType({
            key: _freeze_json(item)
            for key, item in value.items()
        })
    if isinstance(value, (list, tuple)):
        return tuple(_freeze_json(item) for item in value)
    return value


def _thaw_json(value):
    if isinstance(value, Mapping):
        return {
            key: _thaw_json(item)
            for key, item in value.items()
        }
    if isinstance(value, tuple):
        return [_thaw_json(item) for item in value]
    return value


def _utf16_length(value: str) -> int:
    try:
        return len(value.encode("utf-16-le")) // 2
    except UnicodeEncodeError:
        return 1 << 60


def _nested_utf16_length(value) -> int:
    if isinstance(value, str):
        return _utf16_length(value)
    if isinstance(value, Mapping):
        return sum(_nested_utf16_length(item) for item in value.values())
    if isinstance(value, (list, tuple)):
        return sum(_nested_utf16_length(item) for item in value)
    return 0


def _nested_text_field_utf16_length(value) -> int:
    if isinstance(value, Mapping):
        return sum(
            _utf16_length(item)
            if name == "text" and isinstance(item, str)
            else _nested_text_field_utf16_length(item)
            for name, item in value.items()
        )
    if isinstance(value, (list, tuple)):
        return sum(_nested_text_field_utf16_length(item) for item in value)
    return 0


def _json_equal(left, right) -> bool:
    if isinstance(left, bool) or isinstance(right, bool):
        return isinstance(left, bool) and isinstance(right, bool) and left == right
    if left is None or right is None:
        return left is None and right is None
    if isinstance(left, (int, float)) and isinstance(right, (int, float)):
        return left == right
    if isinstance(left, Mapping) and isinstance(right, Mapping):
        return (
            set(left) == set(right)
            and all(_json_equal(left[key], right[key]) for key in left)
        )
    if isinstance(left, (list, tuple)) and isinstance(right, (list, tuple)):
        return len(left) == len(right) and all(
            _json_equal(left_item, right_item)
            for left_item, right_item in zip(left, right)
        )
    return type(left) is type(right) and left == right


def _instance_error(
    schema: Mapping[str, Any],
    value: Any,
    path: str,
    format_validators: Optional[Mapping[str, Callable[[str], bool]]] = None,
) -> Optional[str]:
    if "oneOf" in schema:
        matches = [
            branch
            for branch in schema["oneOf"]
            if _instance_error(
                branch,
                value,
                path,
                format_validators,
            ) is None
        ]
        if len(matches) != 1:
            return f"{path} must match exactly one allowed shape"
        return None
    if "anyOf" in schema:
        if not any(
            _instance_error(
                branch,
                value,
                path,
                format_validators,
            ) is None
            for branch in schema["anyOf"]
        ):
            return f"{path} must match at least one allowed shape"
        return None
    if "const" in schema and not _json_equal(value, schema["const"]):
        return f"{path} must equal {schema['const']!r}"
    expected_type = schema.get("type")
    if expected_type == "object":
        if not isinstance(value, Mapping):
            return f"{path} must be an object"
        properties = schema.get("properties", {})
        missing = [
            name for name in schema.get("required", ())
            if name not in value
        ]
        if missing:
            return f"{path}.{missing[0]} is required"
        if schema.get("additionalProperties") is False:
            unknown = [name for name in value if name not in properties]
            if unknown:
                return f"{path}.{unknown[0]} is not allowed"
        for name, property_schema in properties.items():
            if name in value:
                error = _instance_error(
                    property_schema,
                    value[name],
                    f"{path}.{name}",
                    format_validators,
                )
                if error is not None:
                    return error
        if len(value) < schema.get("minProperties", 0):
            return f"{path} has too few properties"
        if "maxProperties" in schema and len(value) > schema["maxProperties"]:
            return f"{path} has too many properties"
        ordered = schema.get("x-ordered")
        if ordered is not None and all(name in value for name in ordered):
            left, right = ordered
            if value[left] > value[right]:
                return f"{path}.{left} must not exceed {path}.{right}"
        strict_ordered = schema.get("x-strictOrdered")
        if strict_ordered is not None and all(
            name in value for name in strict_ordered
        ):
            left, right = strict_ordered
            if value[left] >= value[right]:
                return f"{path}.{left} must be less than {path}.{right}"
        at_least_one = schema.get("x-atLeastOne")
        if at_least_one is not None and not any(
            name in value for name in at_least_one
        ):
            return f"{path} must contain at least one of {list(at_least_one)}"
        unique_by = schema.get("x-uniqueBy")
        if unique_by is not None:
            seen = set()
            for item in value.get(unique_by[0], ()):
                key = tuple(item.get(name) for name in unique_by[1:])
                if key in seen:
                    return f"{path}.{unique_by[0]} contains duplicate entries"
                seen.add(key)
    elif expected_type == "array":
        if not isinstance(value, (list, tuple)):
            return f"{path} must be an array"
        item_schema = schema.get("items")
        if item_schema is not None:
            for index, item in enumerate(value):
                error = _instance_error(
                    item_schema,
                    item,
                    f"{path}[{index}]",
                    format_validators,
                )
                if error is not None:
                    return error
        if len(value) < schema.get("minItems", 0):
            return f"{path} contains too few items"
        if "maxItems" in schema and len(value) > schema["maxItems"]:
            return f"{path} contains too many items"
        if schema.get("uniqueItems") is True:
            for index, item in enumerate(value):
                if any(
                    _json_equal(item, previous)
                    for previous in value[:index]
                ):
                    return f"{path} items must be unique"
        if schema.get("x-rectangular") is True and value:
            widths = {
                len(row)
                for row in value
                if isinstance(row, (list, tuple))
            }
            if len(widths) != 1:
                return f"{path} must be rectangular"
        if schema.get("x-strictlyIncreasing") is True and any(
            left >= right for left, right in zip(value, value[1:])
        ):
            return f"{path} must be strictly increasing"
        if "x-maxCells" in schema and sum(
            len(row) for row in value if isinstance(row, (list, tuple))
        ) > schema["x-maxCells"]:
            return f"{path} contains too many cells"
    elif expected_type == "string":
        if not isinstance(value, str):
            return f"{path} must be a string"
        if len(value) < schema.get("minLength", 0):
            return f"{path} is too short"
        if "maxLength" in schema and len(value) > schema["maxLength"]:
            return f"{path} is too long"
        if "pattern" in schema and re.search(schema["pattern"], value) is None:
            return f"{path} has an invalid format"
        if "format" in schema and not _matches_format(
            value,
            schema["format"],
            format_validators,
        ):
            return f"{path} must be a valid {schema['format']}"
    elif expected_type == "integer" and (
        not isinstance(value, int) or isinstance(value, bool)
    ):
        return f"{path} must be an integer"
    elif expected_type == "number" and (
        not isinstance(value, (int, float)) or isinstance(value, bool)
    ):
        return f"{path} must be a number"
    elif expected_type == "boolean" and not isinstance(value, bool):
        return f"{path} must be a boolean"
    elif expected_type == "null" and value is not None:
        return f"{path} must be null"
    if expected_type in {"integer", "number"} and isinstance(
        value, (int, float)
    ) and not isinstance(value, bool):
        if not math.isfinite(value):
            return f"{path} must be finite"
        if "minimum" in schema and value < schema["minimum"]:
            return f"{path} must be at least {schema['minimum']}"
        if "maximum" in schema and value > schema["maximum"]:
            return f"{path} must be at most {schema['maximum']}"
    if "enum" in schema and not any(
        _json_equal(value, candidate) for candidate in schema["enum"]
    ):
        return f"{path} must be one of {schema['enum']}"
    if "x-maxUtf16Length" in schema and _nested_utf16_length(
        value
    ) > schema["x-maxUtf16Length"]:
        return f"{path} exceeds its UTF-16 text limit"
    if "x-maxUtf16Text" in schema and _nested_text_field_utf16_length(
        value
    ) > schema["x-maxUtf16Text"]:
        return f"{path} exceeds its UTF-16 payload-text limit"
    return None


def _matches_format(
    value: str,
    format_name: str,
    format_validators: Optional[Mapping[str, Callable[[str], bool]]],
) -> bool:
    validator = (
        None
        if format_validators is None
        else format_validators.get(format_name)
    )
    if validator is None:
        return False
    try:
        return validator(value) is True
    except Exception:
        return False


@dataclass(frozen=True)
class ActionAddress:
    app: str
    action: str

    def __post_init__(self) -> None:
        if self.app not in _APPLICATIONS or not self.action:
            raise ValueError("invalid Action Address")


@dataclass(frozen=True)
class ActionRequest:
    address: ActionAddress
    params: Mapping[str, Any]

    def __post_init__(self) -> None:
        if not isinstance(self.params, Mapping):
            raise ValueError("Action Request params must be an object")


@dataclass(frozen=True)
class TraceContext:
    trace_id: str
    trace_log: Optional[Path]
    event_sink: Optional[Callable[..., None]] = None

    def event(self, name: str, **fields: Any) -> None:
        if self.event_sink is None:
            return
        try:
            self.event_sink(name, **fields)
        except Exception:
            # Diagnostics must never change an Action outcome or binding state.
            pass


@dataclass(frozen=True)
class ControllerContext:
    request_id: str
    trace_id: str
    deadline_at: float


@dataclass(frozen=True)
class ControllerCommand:
    address: ActionAddress
    params: Mapping[str, Any]
    context: ControllerContext


@dataclass(frozen=True)
class ActionContract:
    name: str
    binding_role: str
    risk: str
    parameters: Mapping[str, Any]
    result: Mapping[str, Any]
    purpose: str = ""
    category: str = ""
    stable_errors: Sequence[str] = ()
    verification: str = ""
    prerequisites: Sequence[str] = ()
    constraints: Sequence[str] = ()
    examples: Sequence[Mapping[str, Any]] = ()
    parameter_validator: Optional[Callable[..., Optional[str]]] = None
    semantic_validator: Optional[Callable[..., Optional[str]]] = None

    def __post_init__(self) -> None:
        object.__setattr__(self, "parameters", _freeze_json(self.parameters))
        object.__setattr__(self, "result", _freeze_json(self.result))
        object.__setattr__(self, "stable_errors", tuple(self.stable_errors))
        object.__setattr__(self, "prerequisites", tuple(self.prerequisites))
        object.__setattr__(self, "constraints", tuple(self.constraints))
        object.__setattr__(
            self,
            "examples",
            tuple(_freeze_json(example) for example in self.examples),
        )

    @property
    def required_parameters(self):
        direct = tuple(self.parameters.get("required", ()))
        if direct or "oneOf" not in self.parameters:
            return direct
        branches = tuple(self.parameters["oneOf"])
        if not branches:
            return ()
        common = set(branches[0].get("required", ()))
        for branch in branches[1:]:
            common.intersection_update(branch.get("required", ()))
        return tuple(
            name
            for name in branches[0].get("required", ())
            if name in common
        )

    def to_wire(self, application: str):
        return {
            "address": {"app": application, "action": self.name},
            "purpose": self.purpose,
            "category": self.category,
            "bindingRole": self.binding_role,
            "risk": self.risk,
            "parameters": _thaw_json(self.parameters),
            "result": _thaw_json(self.result),
            "errors": list(self.stable_errors),
            "verification": self.verification,
            "prerequisites": list(self.prerequisites),
            "constraints": list(self.constraints),
            "examples": _thaw_json(self.examples),
        }


@dataclass(frozen=True)
class ActionIndexEntry:
    action: str
    purpose: str
    category: str
    binding_role: str
    risk: str
    required_parameters: Sequence[str]

    def to_wire(self):
        return {
            "action": self.action,
            "purpose": self.purpose,
            "category": self.category,
            "bindingRole": self.binding_role,
            "risk": self.risk,
            "requiredParams": list(self.required_parameters),
        }


_SCHEMA_KEYWORDS = frozenset({
    "type",
    "properties",
    "required",
    "additionalProperties",
    "oneOf",
    "anyOf",
    "const",
    "items",
    "minItems",
    "maxItems",
    "uniqueItems",
    "minLength",
    "maxLength",
    "pattern",
    "format",
    "enum",
    "minimum",
    "maximum",
    "minProperties",
    "maxProperties",
    "x-ordered",
    "x-strictOrdered",
    "x-atLeastOne",
    "x-uniqueBy",
    "x-rectangular",
    "x-strictlyIncreasing",
    "x-maxCells",
    "x-maxUtf16Length",
    "x-maxUtf16Text",
})

_SCHEMA_TYPES = frozenset({
    "object",
    "array",
    "string",
    "integer",
    "number",
    "boolean",
    "null",
})

def _schema_definition_error(schema, path, registered_formats=frozenset()):
    if not isinstance(schema, Mapping):
        return f"{path} must be an object schema"
    unknown = set(schema) - _SCHEMA_KEYWORDS
    if unknown:
        return f"{path} contains unknown keyword {sorted(unknown)[0]}"
    if "type" in schema and schema["type"] not in _SCHEMA_TYPES:
        return f"{path}.type is invalid"
    if "format" in schema and schema["format"] not in registered_formats:
        return f"{path}.format is not registered"
    for keyword in ("oneOf", "anyOf"):
        if keyword in schema:
            branches = schema[keyword]
            if not isinstance(branches, (list, tuple)) or not branches:
                return f"{path}.{keyword} must be non-empty"
            for index, branch in enumerate(branches):
                error = _schema_definition_error(
                    branch,
                    f"{path}.{keyword}[{index}]",
                    registered_formats,
                )
                if error is not None:
                    return error
    if schema.get("type") == "object":
        properties = schema.get("properties", {})
        required = schema.get("required", ())
        if not isinstance(properties, Mapping):
            return f"{path}.properties must be an object"
        if len(required) != len(set(required)):
            return f"{path}.required contains duplicates"
        if not set(required).issubset(properties):
            return f"{path}.required names an unknown property"
        for name, child in properties.items():
            error = _schema_definition_error(
                child,
                f"{path}.{name}",
                registered_formats,
            )
            if error is not None:
                return error
    if schema.get("type") == "array":
        if "items" not in schema:
            return f"{path}.items is required"
        error = _schema_definition_error(
            schema["items"],
            f"{path}.items",
            registered_formats,
        )
        if error is not None:
            return error
    for minimum, maximum in (
        ("minLength", "maxLength"),
        ("minItems", "maxItems"),
        ("minProperties", "maxProperties"),
        ("minimum", "maximum"),
    ):
        if (
            minimum in schema
            and maximum in schema
            and schema[minimum] > schema[maximum]
        ):
            return f"{path}.{minimum} exceeds {maximum}"
    return None


class ApplicationContractSet:
    def __init__(
        self,
        *,
        application: str,
        contracts: Sequence[ActionContract],
        format_validators: Optional[
            Mapping[str, Callable[[str], bool]]
        ] = None,
        allow_incomplete: bool = False,
    ):
        if application not in _APPLICATIONS:
            raise ValueError(f"invalid application: {application}")
        if format_validators is None:
            format_validators = {}
        if not isinstance(format_validators, Mapping) or any(
            not isinstance(name, str)
            or not name
            or not callable(validator)
            for name, validator in format_validators.items()
        ):
            raise ValueError("invalid application format validators")
        self._format_validators = MappingProxyType(dict(format_validators))
        registered_formats = frozenset(self._format_validators)
        names = [contract.name for contract in contracts]
        if len(names) != len(set(names)):
            raise ValueError(
                "Application Contract Set contains duplicate Action names"
            )
        for contract in contracts:
            if not contract.name:
                raise ValueError("Action name must be non-empty")
            if contract.binding_role not in _BINDING_ROLES:
                raise ValueError(
                    f"invalid Document Binding Role for {contract.name}"
                )
            if contract.risk not in _RISKS:
                raise ValueError(f"invalid risk for {contract.name}")
            if not allow_incomplete:
                if not contract.purpose:
                    raise ValueError(f"missing purpose for {contract.name}")
                if not contract.category:
                    raise ValueError(f"missing category for {contract.name}")
                if not contract.verification:
                    raise ValueError(
                        f"missing verification guidance for {contract.name}"
                    )
                if not contract.prerequisites:
                    raise ValueError(
                        f"missing prerequisites for {contract.name}"
                    )
                if not contract.constraints:
                    raise ValueError(
                        f"missing constraints for {contract.name}"
                    )
                if not contract.examples:
                    raise ValueError(f"missing examples for {contract.name}")
                if not contract.stable_errors:
                    raise ValueError(
                        f"missing stable errors for {contract.name}"
                    )
                if (
                    contract.parameter_validator is not None
                    and not callable(contract.parameter_validator)
                ):
                    raise ValueError(
                        f"invalid parameter validator for {contract.name}"
                    )
                if (
                    contract.semantic_validator is not None
                    and not callable(contract.semantic_validator)
                ):
                    raise ValueError(
                        f"invalid semantic validator for {contract.name}"
                    )
                if len(contract.stable_errors) != len(
                    set(contract.stable_errors)
                ) or any(
                    re.fullmatch(r"[A-Z][A-Z0-9_]*", code) is None
                    for code in contract.stable_errors
                ):
                    raise ValueError(
                        f"invalid stable errors for {contract.name}"
                    )
                for location, schema in (
                    ("parameters", contract.parameters),
                    ("result", contract.result),
                ):
                    error = _schema_definition_error(
                        schema,
                        f"{contract.name}.{location}",
                        registered_formats,
                    )
                    if error is not None:
                        raise ValueError(error)
                for index, example in enumerate(contract.examples):
                    if not isinstance(example, Mapping) or set(example) != {
                        "params"
                    }:
                        raise ValueError(
                            f"invalid example {index} for {contract.name}"
                        )
                    error = _instance_error(
                        contract.parameters,
                        example["params"],
                        f"{contract.name}.examples[{index}].params",
                        self._format_validators,
                    )
                    if error is not None:
                        raise ValueError(error)
                    if contract.parameter_validator is not None:
                        semantic_error = contract.parameter_validator(
                            example["params"]
                        )
                        if semantic_error is not None:
                            raise ValueError(
                                f"{contract.name}.examples[{index}].params: "
                                f"{semantic_error}"
                            )
        self.application = application
        self._ordered_contracts = tuple(contracts)
        self._contracts = {
            contract.name: contract
            for contract in contracts
        }

    def resolve(self, action: str) -> Optional[ActionContract]:
        return self._contracts.get(action)

    @property
    def action_names(self):
        return tuple(contract.name for contract in self._ordered_contracts)

    @property
    def contracts(self):
        return self._ordered_contracts

    def action_index(self):
        return tuple(
            ActionIndexEntry(
                action=contract.name,
                purpose=contract.purpose,
                category=contract.category,
                binding_role=contract.binding_role,
                risk=contract.risk,
                required_parameters=contract.required_parameters,
            )
            for contract in self._ordered_contracts
        )

    def validate_params(self, action: str, params: Mapping[str, Any]) -> None:
        contract = self.resolve(action)
        if contract is None:
            raise ContractValidationError(
                action=action,
                location="params",
                message="unknown Action",
            )
        error = _instance_error(
            contract.parameters,
            params,
            "params",
            self._format_validators,
        )
        if error is not None:
            raise ContractValidationError(
                action=action,
                location="params",
                message=error,
            )
        if contract.parameter_validator is not None:
            semantic_error = contract.parameter_validator(params)
            if semantic_error is not None:
                raise ContractValidationError(
                    action=action,
                    location="params",
                    message=semantic_error,
                )

    def validate_result(
        self,
        action: str,
        result: Mapping[str, Any],
        *,
        params: Optional[Mapping[str, Any]] = None,
    ) -> None:
        contract = self.resolve(action)
        if contract is None:
            raise ContractValidationError(
                action=action,
                location="result",
                message="unknown Action",
            )
        error = _instance_error(
            contract.result,
            result,
            "result",
            self._format_validators,
        )
        if error is not None:
            raise ContractValidationError(
                action=action,
                location="result",
                message=error,
            )
        if contract.semantic_validator is not None and params is None:
            raise ContractValidationError(
                action=action,
                location="result",
                message=(
                    "the original params are required for semantic result "
                    "validation"
                ),
            )
        if params is not None:
            self.validate_params(action, params)
        if contract.semantic_validator is not None:
            semantic_error = contract.semantic_validator(params, result)
            if semantic_error is not None:
                raise ContractValidationError(
                    action=action,
                    location="result",
                    message=semantic_error,
                )

    def resolve_wire(self, action: str):
        contract = self.resolve(action)
        if contract is None:
            raise ContractValidationError(
                action=action,
                location="contract",
                message="unknown Action",
            )
        return contract.to_wire(self.application)

    def batch_resolve(self, addresses: Sequence[ActionAddress]):
        items = []
        resolved_count = 0
        for address in addresses:
            item = {
                "address": {
                    "app": address.app,
                    "action": address.action,
                }
            }
            if address.app != self.application:
                item.update({
                    "status": "unresolved",
                    "error": {
                        "code": "SESSION_APP_MISMATCH",
                        "message": (
                            f"Contract Set application is {self.application}, "
                            f"not {address.app}"
                        ),
                    },
                })
            else:
                contract = self.resolve(address.action)
                if contract is None:
                    item.update({
                        "status": "unresolved",
                        "error": {
                            "code": "UNKNOWN_ACTION",
                            "message": (
                                f"Unknown {self.application} Action: "
                                f"{address.action}"
                            ),
                        },
                    })
                else:
                    resolved_count += 1
                    item.update({
                        "status": "resolved",
                        "contract": contract.to_wire(self.application),
                    })
            items.append(item)
        if resolved_count == len(items) and items:
            status = "complete"
        elif resolved_count:
            status = "partial"
        else:
            status = "failed"
        return {"status": status, "items": items}


@dataclass(frozen=True)
class ActionError:
    code: str
    message: str

    def __post_init__(self) -> None:
        if not isinstance(self.code, str) or not self.code:
            raise ValueError("Action error code must be non-empty")
        if not isinstance(self.message, str) or not self.message:
            raise ValueError("Action error message must be non-empty")


@dataclass(frozen=True)
class ControllerResult:
    outcome: str
    controller_state: str
    binding_disposition: str
    data: Optional[Mapping[str, Any]] = None
    error: Optional[ActionError] = None

    def __post_init__(self) -> None:
        valid_tags = (
            self.outcome in _OUTCOMES
            and self.controller_state in _CONTROLLER_STATES
            and self.binding_disposition in {
                "unchanged",
                "established",
                "not_established",
                "lost",
                "unprovable",
            }
        )
        valid_variant = (
            isinstance(self.data, Mapping) and self.error is None
            if self.outcome == "succeeded"
            else isinstance(self.error, ActionError) and self.data is None
        )
        if not valid_tags or not valid_variant:
            raise ValueError("ControllerResult must be one closed variant")
        if self.data is not None:
            object.__setattr__(self, "data", _freeze_json(self.data))

    @classmethod
    def succeeded(
        cls,
        *,
        data: Mapping[str, Any],
        controller_state: str,
        binding_disposition: str,
    ) -> "ControllerResult":
        return cls(
            outcome="succeeded",
            data=data,
            controller_state=controller_state,
            binding_disposition=binding_disposition,
        )

    @classmethod
    def failed(
        cls,
        *,
        error: ActionError,
        controller_state: str,
        binding_disposition: str,
    ) -> "ControllerResult":
        return cls(
            outcome="failed",
            error=error,
            controller_state=controller_state,
            binding_disposition=binding_disposition,
        )

    @classmethod
    def unknown(
        cls,
        *,
        error: ActionError,
        controller_state: str,
        binding_disposition: str,
    ) -> "ControllerResult":
        return cls(
            outcome="unknown",
            error=error,
            controller_state=controller_state,
            binding_disposition=binding_disposition,
        )


class _ControlPlaneControllerResult(ControllerResult):
    """Controller-owned lifecycle result that an Adapter cannot impersonate."""


@dataclass(frozen=True)
class ActionResponse:
    outcome: str
    address: ActionAddress
    session_id: str
    trace: TraceContext
    error: Optional[ActionError] = None
    data: Optional[Mapping[str, Any]] = None

    def __post_init__(self) -> None:
        if self.outcome == "succeeded":
            valid = isinstance(self.data, Mapping) and self.error is None
        elif self.outcome in {"failed", "unknown"}:
            valid = isinstance(self.error, ActionError) and self.data is None
        else:
            valid = False
        if not valid:
            raise ValueError("ActionResponse must be one exact closed variant")
        if self.data is not None:
            object.__setattr__(self, "data", _freeze_json(self.data))

    def to_wire(self) -> Mapping[str, Any]:
        wire = {
            "outcome": self.outcome,
            "address": {
                "app": self.address.app,
                "action": self.address.action,
            },
        }
        if self.outcome == "succeeded":
            wire["data"] = _thaw_json(self.data)
        else:
            wire["error"] = {
                "code": self.error.code,
                "message": self.error.message,
            }
        wire.update({
            "sessionId": self.session_id,
            "traceId": self.trace.trace_id,
            "traceLog": (
                str(self.trace.trace_log)
                if self.trace.trace_log is not None
                else None
            ),
        })
        return wire


@dataclass(frozen=True)
class ActionTurn:
    response: ActionResponse
    continuation: str

    def __post_init__(self) -> None:
        if self.continuation not in {"continue", "terminate"}:
            raise ValueError("invalid Action continuation")


@dataclass(frozen=True)
class _RuntimeDisposition:
    outcome: str
    address: ActionAddress
    trace: TraceContext
    error: Optional[ActionError] = None
    data: Optional[Mapping[str, Any]] = None

    def __post_init__(self) -> None:
        if self.outcome == "succeeded":
            valid = isinstance(self.data, Mapping) and self.error is None
        else:
            valid = (
                self.outcome in {"failed", "unknown"}
                and isinstance(self.error, ActionError)
                and self.data is None
            )
        if not valid:
            raise ValueError("invalid Runtime disposition")


@dataclass(frozen=True)
class _RuntimeTurn:
    disposition: _RuntimeDisposition
    continuation: str

    def __post_init__(self) -> None:
        if self.continuation not in {"continue", "terminate"}:
            raise ValueError("invalid Runtime continuation")


@dataclass(frozen=True)
class AcquiredDocument:
    document: Any
    data: Mapping[str, Any]


@dataclass(frozen=True)
class PreparedDocumentAcquisition:
    """Application-owned plan whose identity can be guarded before effects."""

    coordination_identity: str
    application_state: Any

    def __post_init__(self) -> None:
        if (
            not isinstance(self.coordination_identity, str)
            or not self.coordination_identity
        ):
            raise ValueError("coordination identity must be non-empty")


class DefiniteEstablishFailure(Exception):
    def __init__(self, *, code: str, message: str):
        super().__init__(message)
        self.code = code
        self.message = message


class UnprovableEstablishFailure(Exception):
    def __init__(
        self,
        *,
        outcome: str,
        code: str,
        message: str,
        partial_document: Any = None,
    ):
        super().__init__(message)
        self.outcome = outcome
        self.code = code
        self.message = message
        self.partial_document = partial_document


class RequiredBindingFailure(Exception):
    def __init__(
        self,
        *,
        disposition: str,
        outcome: str,
        code: str,
        message: str,
    ):
        super().__init__(message)
        self.disposition = disposition
        self.outcome = outcome
        self.code = code
        self.message = message


@runtime_checkable
class ApplicationAdapter(Protocol):
    """Core-facing seam implemented by one application-local Adapter."""

    @property
    def application(self) -> str:
        ...

    def prepare_establish(
        self,
        command: ControllerCommand,
    ) -> PreparedDocumentAcquisition:
        ...

    def establish(
        self,
        prepared: PreparedDocumentAcquisition,
        command: ControllerCommand,
    ) -> AcquiredDocument:
        ...

    def is_live(self, document: Any) -> bool:
        ...

    def handle(
        self,
        document: Any,
        command: ControllerCommand,
    ) -> ControllerResult:
        ...

    def handle_none(self, command: ControllerCommand) -> ControllerResult:
        ...

    def close(self) -> bool:
        ...


@dataclass(frozen=True)
class DocumentResourceCleanup:
    state: str

    def __post_init__(self) -> None:
        if self.state not in {
            "not_acquired",
            "released",
            "quarantined",
            "release_unconfirmed",
        }:
            raise ValueError("invalid document cleanup state")


@dataclass(frozen=True)
class ProcessCleanup:
    pid: int
    cleanup_steps: Sequence[str]
    released: bool

    def __post_init__(self) -> None:
        if (
            not isinstance(self.pid, int)
            or isinstance(self.pid, bool)
            or self.pid < 1
        ):
            raise ValueError("cleanup process pid must be positive")
        steps = tuple(self.cleanup_steps)
        if not steps or any(
            step not in {
                "already_exited",
                "graceful",
                "terminate",
                "kill",
                "job_close",
            }
            for step in steps
        ):
            raise ValueError("invalid process cleanup steps")
        if not isinstance(self.released, bool):
            raise ValueError("process cleanup released must be Boolean")
        object.__setattr__(self, "cleanup_steps", steps)


@dataclass(frozen=True)
class CleanupReport:
    document_resources: DocumentResourceCleanup
    processes: Sequence[ProcessCleanup] = ()

    def __post_init__(self) -> None:
        processes = tuple(self.processes)
        if any(not isinstance(item, ProcessCleanup) for item in processes):
            raise ValueError("cleanup processes must be Process Cleanup values")
        object.__setattr__(self, "processes", processes)


@dataclass(frozen=True)
class CleanupOutcome:
    outcome: str
    cleanup: CleanupReport
    error: Optional[ActionError] = None

    def __post_init__(self) -> None:
        valid = (
            self.outcome == "succeeded" and self.error is None
        ) or (
            self.outcome == "failed" and isinstance(self.error, ActionError)
        )
        if not valid:
            raise ValueError("CleanupOutcome must be one closed variant")


class _BindingController:
    """Own the exact document, partial acquisition state, and Lease."""

    def __init__(
        self,
        *,
        contracts,
        adapter,
        coordinator,
        cleanup_timeout_seconds,
    ):
        self._contracts = contracts
        self._adapter = adapter
        self._coordinator = coordinator
        self._cleanup_timeout_seconds = cleanup_timeout_seconds
        self._guard = None
        self._document = None
        self._lease = None
        self._cleanup = None
        self._state_lock = threading.Lock()
        self._close_lock = threading.Lock()
        self._in_flight = False
        self._closing = False
        self._release_in_progress = False

    def _is_closing(self):
        with self._state_lock:
            return self._closing

    @staticmethod
    def _runtime_closed_result(binding_disposition):
        return _ControlPlaneControllerResult(
            outcome="failed",
            error=ActionError(
                code="RUNTIME_CLOSED",
                message="The controller is closing",
            ),
            controller_state="broken",
            binding_disposition=binding_disposition,
        )

    @staticmethod
    def _error_result(
        *,
        outcome,
        code,
        message,
        controller_state,
        binding_disposition,
    ):
        return ControllerResult(
            outcome=outcome,
            error=ActionError(code=code, message=message),
            controller_state=controller_state,
            binding_disposition=binding_disposition,
        )

    def _establish(self, command):
        try:
            prepared = self._adapter.prepare_establish(command)
            if not isinstance(prepared, PreparedDocumentAcquisition):
                raise TypeError(
                    "The Adapter returned an invalid acquisition plan"
                )
            guard = self._coordinator.begin(
                prepared.coordination_identity,
                command.context,
            )
            with self._state_lock:
                if self._closing:
                    return self._error_result(
                        outcome="unknown",
                        code="RESPONSE_LOST",
                        message="Session cleanup began during acquisition",
                        controller_state="broken",
                        binding_disposition="unprovable",
                    )
                self._guard = guard
            acquired = self._adapter.establish(prepared, command)
        except DefiniteEstablishFailure as exc:
            with self._state_lock:
                if self._closing:
                    return self._error_result(
                        outcome="failed",
                        code=exc.code,
                        message=exc.message,
                        controller_state="broken",
                        binding_disposition="unprovable",
                    )
                guard = self._guard
                self._guard = None
                self._release_in_progress = guard is not None
            try:
                released = DocumentResourceCleanup(state="not_acquired")
                if guard is not None:
                    released = self._coordinator.release(
                        guard,
                        None,
                        None,
                        in_flight=False,
                    )
            except Exception:
                released = DocumentResourceCleanup(
                    state="release_unconfirmed"
                )
            with self._state_lock:
                self._release_in_progress = False
                cleanup_started = self._cleanup is not None
            if cleanup_started:
                return self._error_result(
                    outcome="failed",
                    code=exc.code,
                    message=exc.message,
                    controller_state="broken",
                    binding_disposition="unprovable",
                )
            if (
                not isinstance(released, DocumentResourceCleanup)
                or released.state not in {"not_acquired", "released"}
            ):
                with self._state_lock:
                    self._cleanup = (
                        released
                        if isinstance(released, DocumentResourceCleanup)
                        else DocumentResourceCleanup(
                            state="release_unconfirmed"
                        )
                    )
                return self._error_result(
                    outcome="failed",
                    code=exc.code,
                    message=exc.message,
                    controller_state="broken",
                    binding_disposition="unprovable",
                )
            return self._error_result(
                outcome="failed",
                code=exc.code,
                message=exc.message,
                controller_state="usable",
                binding_disposition="not_established",
            )
        except UnprovableEstablishFailure as exc:
            with self._state_lock:
                if not self._closing:
                    self._document = exc.partial_document
            return self._error_result(
                outcome=exc.outcome,
                code=exc.code,
                message=exc.message,
                controller_state="broken",
                binding_disposition="unprovable",
            )
        except Exception:
            return self._error_result(
                outcome="unknown",
                code="RESPONSE_LOST",
                message="No reliable establish result was received",
                controller_state="broken",
                binding_disposition="unprovable",
            )
        if not isinstance(acquired, AcquiredDocument):
            return self._error_result(
                outcome="unknown",
                code="INVALID_RESULT",
                message="The Adapter returned an invalid document acquisition",
                controller_state="broken",
                binding_disposition="unprovable",
            )
        if self._is_closing():
            return self._error_result(
                outcome="unknown",
                code="RESPONSE_LOST",
                message="Session cleanup began during document acquisition",
                controller_state="broken",
                binding_disposition="unprovable",
            )
        with self._state_lock:
            if self._closing:
                return self._error_result(
                    outcome="unknown",
                    code="RESPONSE_LOST",
                    message="Session cleanup began during document acquisition",
                    controller_state="broken",
                    binding_disposition="unprovable",
                )
            self._document = acquired.document
            guard = self._guard
        try:
            lease = self._coordinator.commit(
                guard,
                acquired.document,
                command.context,
            )
        except Exception:
            return self._error_result(
                outcome="failed",
                code="DOCUMENT_BINDING_UNAVAILABLE",
                message=(
                    "The document and Lease could not be committed "
                    "as one binding"
                ),
                controller_state="broken",
                binding_disposition="unprovable",
            )
        with self._state_lock:
            if self._closing:
                self._guard = None
                self._document = None
                self._lease = None
                return self._error_result(
                    outcome="unknown",
                    code="RESPONSE_LOST",
                    message="Session cleanup began during Lease commit",
                    controller_state="broken",
                    binding_disposition="unprovable",
                )
            self._lease = lease
            self._guard = None
        return ControllerResult.succeeded(
            data=acquired.data,
            controller_state="usable",
            binding_disposition="established",
        )

    def _required(self, command):
        with self._state_lock:
            if self._closing:
                return self._runtime_closed_result("unprovable")
            document = self._document
        try:
            live = self._adapter.is_live(document)
        except Exception:
            return self._error_result(
                outcome="failed",
                code="DOCUMENT_BINDING_UNAVAILABLE",
                message="The bound document liveness could not be proved",
                controller_state="broken",
                binding_disposition="unprovable",
            )
        if not live:
            return self._error_result(
                outcome="failed",
                code="DOCUMENT_CLOSED",
                message="The bound document is closed",
                controller_state="broken",
                binding_disposition="lost",
            )
        try:
            return self._adapter.handle(document, command)
        except RequiredBindingFailure as exc:
            return self._error_result(
                outcome=exc.outcome,
                code=exc.code,
                message=exc.message,
                controller_state="broken",
                binding_disposition=exc.disposition,
            )
        except Exception:
            return self._error_result(
                outcome="unknown",
                code="RESPONSE_LOST",
                message="The controller did not return a reliable result",
                controller_state="broken",
                binding_disposition="unprovable",
            )

    def _none(self, command):
        try:
            return self._adapter.handle_none(command)
        except Exception:
            return self._error_result(
                outcome="unknown",
                code="RESPONSE_LOST",
                message="The controller did not return a reliable result",
                controller_state="broken",
                binding_disposition="unchanged",
            )

    def execute(self, command):
        with self._state_lock:
            if self._closing:
                contract = self._contracts.resolve(command.address.action)
                return self._runtime_closed_result(
                    "unchanged"
                    if contract is not None and contract.binding_role == "none"
                    else "unprovable"
                )
            self._in_flight = True
        try:
            contract = self._contracts.resolve(command.address.action)
            if contract.binding_role == "establish":
                return self._establish(command)
            if contract.binding_role == "required":
                return self._required(command)
            return self._none(command)
        finally:
            with self._state_lock:
                self._in_flight = False

    def close(self):
        with self._close_lock:
            if self._cleanup is not None:
                return self._cleanup
            with self._state_lock:
                self._closing = True
                in_flight = self._in_flight
                guard = self._guard
                document = self._document
                lease = self._lease
                release_in_progress = self._release_in_progress
            if release_in_progress:
                self._cleanup = DocumentResourceCleanup(
                    state="release_unconfirmed"
                )
            elif in_flight or any(
                resource is not None
                for resource in (guard, document, lease)
            ):
                released = {}

                def release_resources():
                    try:
                        released["value"] = self._coordinator.release(
                            guard,
                            document,
                            lease,
                            in_flight=in_flight,
                        )
                    except BaseException as exc:
                        released["error"] = exc

                worker = threading.Thread(
                    target=release_resources,
                    daemon=True,
                )
                worker.start()
                worker.join(timeout=self._cleanup_timeout_seconds)
                if (
                    worker.is_alive()
                    or "error" in released
                    or not isinstance(
                        released.get("value"),
                        DocumentResourceCleanup,
                    )
                ):
                    self._cleanup = DocumentResourceCleanup(
                        state="release_unconfirmed"
                    )
                else:
                    self._cleanup = released["value"]
            else:
                self._cleanup = DocumentResourceCleanup(
                    state="not_acquired"
                )
            with self._state_lock:
                self._guard = None
                self._document = None
                self._lease = None
            return self._cleanup


class _ActionRuntime:
    """Own contract resolution, logical binding state, and one controller."""

    def __init__(
        self,
        *,
        application: str,
        contracts: ApplicationContractSet,
        adapter: ApplicationAdapter,
        coordinator: Any,
        request_id_factory: Optional[Callable[[], str]] = None,
        clock: Optional[Callable[[], float]] = None,
        action_timeout_seconds: float = 120,
        cleanup_timeout_seconds: float = 5,
    ):
        if contracts.application != application:
            raise ValueError(
                "Contract Set application must match the Session"
            )
        if not isinstance(adapter, ApplicationAdapter):
            raise ValueError("Adapter must satisfy the Application Adapter seam")
        if adapter.application != application:
            raise ValueError("Adapter application must match the Session")
        self._application = application
        self._contracts = contracts
        self._controller_factory = lambda: _BindingController(
            contracts=contracts,
            adapter=adapter,
            coordinator=coordinator,
            cleanup_timeout_seconds=cleanup_timeout_seconds,
        )
        self._controller = None
        self._request_id_factory = (
            request_id_factory
            or (lambda: uuid.uuid4().hex)
        )
        self._clock = clock or time.monotonic
        self._action_timeout_seconds = action_timeout_seconds
        self._binding_state = "UNBOUND"
        self._terminal = False
        self._closed = False
        self._cleanup_outcome = None
        self._execute_lock = threading.Lock()
        self._close_lock = threading.Lock()
        self._state_lock = threading.Lock()

    def _controller_for_dispatch(self):
        with self._state_lock:
            if self._closed or self._terminal:
                return None
            if self._controller is None:
                self._controller = self._controller_factory()
            return self._controller

    def _turn(
        self,
        disposition: _RuntimeDisposition,
        continuation: str,
    ) -> _RuntimeTurn:
        if continuation == "terminate":
            with self._state_lock:
                self._terminal = True
        return _RuntimeTurn(
            disposition=disposition,
            continuation=continuation,
        )

    def _error_turn(
        self,
        *,
        request: ActionRequest,
        trace: TraceContext,
        code: str,
        message: str,
        outcome: str = "failed",
        continuation: str = "continue",
    ) -> _RuntimeTurn:
        return self._turn(
            _RuntimeDisposition(
                outcome=outcome,
                address=request.address,
                trace=trace,
                error=ActionError(code=code, message=message),
            ),
            continuation,
        )

    def _success_turn(
        self,
        *,
        request: ActionRequest,
        trace: TraceContext,
        data: Mapping[str, Any],
        continuation: str,
    ) -> _RuntimeTurn:
        return self._turn(
            _RuntimeDisposition(
                outcome="succeeded",
                address=request.address,
                trace=trace,
                data=data,
            ),
            continuation,
        )

    def _command(
        self,
        request: ActionRequest,
        trace: TraceContext,
    ) -> ControllerCommand:
        return ControllerCommand(
            address=request.address,
            params=_freeze_json(request.params),
            context=ControllerContext(
                request_id=self._request_id_factory(),
                trace_id=trace.trace_id,
                deadline_at=self._clock() + self._action_timeout_seconds,
            ),
        )

    @staticmethod
    def _controller_result_is_valid(
        result: Any,
        *,
        allowed_dispositions,
    ) -> bool:
        if (
            not isinstance(result, ControllerResult)
            or result.outcome not in _OUTCOMES
            or result.controller_state not in _CONTROLLER_STATES
            or result.binding_disposition not in allowed_dispositions
        ):
            return False
        if result.outcome == "succeeded":
            return result.data is not None and result.error is None
        return result.error is not None and result.data is None

    def _map_controller_result(
        self,
        *,
        contract: ActionContract,
        result: Any,
        request: ActionRequest,
        trace: TraceContext,
        allowed_dispositions,
    ) -> _RuntimeTurn:
        if not self._controller_result_is_valid(
            result,
            allowed_dispositions=allowed_dispositions,
        ):
            return self._error_turn(
                request=request,
                trace=trace,
                outcome="failed" if contract.risk == "read" else "unknown",
                code="INVALID_RESULT",
                message="The controller returned an invalid result",
                continuation="terminate",
            )
        if contract.binding_role == "establish":
            if (
                result.outcome == "succeeded"
                and result.binding_disposition == "established"
            ):
                self._binding_state = "BOUND"
            elif (
                result.outcome == "failed"
                and result.binding_disposition == "not_established"
            ):
                self._binding_state = "UNBOUND"
        continuation = (
            "terminate"
            if result.controller_state == "broken"
            or result.binding_disposition in {"lost", "unprovable"}
            or (
                contract.binding_role == "establish"
                and result.outcome == "unknown"
            )
            else "continue"
        )
        if result.outcome == "succeeded":
            try:
                self._contracts.validate_result(
                    contract.name,
                    result.data,
                    params=request.params,
                )
                result_error = None
            except ContractValidationError as exc:
                result_error = exc.message
            if result_error is not None:
                return self._error_turn(
                    request=request,
                    trace=trace,
                    outcome=(
                        "failed" if contract.risk == "read" else "unknown"
                    ),
                    code="INVALID_RESULT",
                    message=result_error,
                    continuation=(
                        "terminate"
                        if contract.binding_role == "establish"
                        else continuation
                    ),
                )
            return self._success_turn(
                request=request,
                trace=trace,
                data=result.data,
                continuation=continuation,
            )
        response_outcome = (
            "failed"
            if result.outcome == "unknown" and contract.risk == "read"
            else result.outcome
        )
        error = result.error
        is_control_plane_error = (
            isinstance(result, _ControlPlaneControllerResult)
            and error.code in _CONTROL_PLANE_ERROR_CODES
            and result.outcome == "failed"
            and result.controller_state == "broken"
            and result.binding_disposition == (
                "unchanged"
                if contract.binding_role == "none"
                else "unprovable"
            )
        )
        if (
            result.outcome == "failed"
            and result.binding_disposition == "unprovable"
            and contract.binding_role in {"establish", "required"}
            and not is_control_plane_error
        ):
            error = ActionError(
                code="DOCUMENT_BINDING_UNAVAILABLE",
                message=result.error.message,
            )
        if (
            contract.stable_errors
            and error.code not in contract.stable_errors
            and not is_control_plane_error
        ):
            return self._error_turn(
                request=request,
                trace=trace,
                outcome="failed" if contract.risk == "read" else "unknown",
                code="INVALID_RESULT",
                message=(
                    f"Controller error code is not declared by {contract.name}"
                ),
                continuation="terminate",
            )
        return self._error_turn(
            request=request,
            trace=trace,
            outcome=response_outcome,
            code=error.code,
            message=error.message,
            continuation=continuation,
        )

    def _execute_none(
        self,
        *,
        controller,
        contract: ActionContract,
        command: ControllerCommand,
        request: ActionRequest,
        trace: TraceContext,
    ) -> _RuntimeTurn:
        result = controller.execute(command)
        return self._map_controller_result(
            contract=contract,
            result=result,
            request=request,
            trace=trace,
            allowed_dispositions={"unchanged"},
        )

    def _execute_establish(
        self,
        *,
        controller,
        contract: ActionContract,
        command: ControllerCommand,
        request: ActionRequest,
        trace: TraceContext,
    ) -> _RuntimeTurn:
        self._binding_state = "ESTABLISHING"
        trace.event("binding.establishing")
        result = controller.execute(command)
        allowed_dispositions = (
            {"established"}
            if isinstance(result, ControllerResult)
            and result.outcome == "succeeded"
            else {"not_established", "unprovable"}
        )
        return self._map_controller_result(
            contract=contract,
            result=result,
            request=request,
            trace=trace,
            allowed_dispositions=allowed_dispositions,
        )

    def _execute_required(
        self,
        *,
        controller,
        contract: ActionContract,
        command: ControllerCommand,
        request: ActionRequest,
        trace: TraceContext,
    ) -> _RuntimeTurn:
        result = controller.execute(command)
        return self._map_controller_result(
            contract=contract,
            result=result,
            request=request,
            trace=trace,
            allowed_dispositions={"unchanged", "lost", "unprovable"},
        )

    def _execute_once(
        self,
        request: ActionRequest,
        trace: TraceContext,
    ) -> _RuntimeTurn:
        with self._state_lock:
            unavailable = self._closed or self._terminal
        if unavailable:
            return self._error_turn(
                request=request,
                trace=trace,
                code="RUNTIME_CLOSED",
                message="The Action Session cannot accept another Action",
                continuation="terminate",
            )
        if request.address.app != self._application:
            return self._error_turn(
                request=request,
                trace=trace,
                code="SESSION_APP_MISMATCH",
                message=(
                    f"The Session application is {self._application}, not "
                    f"{request.address.app}"
                ),
            )
        contract = self._contracts.resolve(request.address.action)
        if contract is None:
            return self._error_turn(
                request=request,
                trace=trace,
                code="UNKNOWN_ACTION",
                message=(
                    f"Unknown {self._application} Action: "
                    f"{request.address.action}"
                ),
            )
        if (
            contract.binding_role == "required"
            and self._binding_state == "UNBOUND"
        ):
            return self._error_turn(
                request=request,
                trace=trace,
                code="SESSION_DOCUMENT_NOT_BOUND",
                message="The Action Session has no bound document",
            )
        if (
            contract.binding_role == "establish"
            and self._binding_state == "BOUND"
        ):
            return self._error_turn(
                request=request,
                trace=trace,
                code="SESSION_DOCUMENT_ALREADY_BOUND",
                message="The Action Session is already bound to a document",
            )
        try:
            self._contracts.validate_params(contract.name, request.params)
            parameter_error = None
        except ContractValidationError as exc:
            parameter_error = exc.message
        if parameter_error is not None:
            return self._error_turn(
                request=request,
                trace=trace,
                code="INVALID_PARAMS",
                message=parameter_error,
            )
        command = self._command(request, trace)
        controller = self._controller_for_dispatch()
        if controller is None:
            return self._error_turn(
                request=request,
                trace=trace,
                code="RUNTIME_CLOSED",
                message="The Action Session cannot accept another Action",
                continuation="terminate",
            )
        if contract.binding_role == "none":
            return self._execute_none(
                controller=controller,
                contract=contract,
                command=command,
                request=request,
                trace=trace,
            )
        if contract.binding_role == "establish":
            return self._execute_establish(
                controller=controller,
                contract=contract,
                command=command,
                request=request,
                trace=trace,
            )
        return self._execute_required(
            controller=controller,
            contract=contract,
            command=command,
            request=request,
            trace=trace,
        )

    def execute(
        self,
        request: ActionRequest,
        trace: TraceContext,
    ) -> _RuntimeTurn:
        with self._execute_lock:
            return self._execute_once(request, trace)

    def _close_once(self) -> CleanupOutcome:
        if self._cleanup_outcome is not None:
            return self._cleanup_outcome
        with self._state_lock:
            self._closed = True
            controller = self._controller
        if controller is None:
            document_resources = DocumentResourceCleanup(
                state="not_acquired"
            )
        else:
            document_resources = controller.close()
        cleanup_succeeded = document_resources.state in {
            "not_acquired",
            "released",
        }
        self._cleanup_outcome = CleanupOutcome(
            outcome="succeeded" if cleanup_succeeded else "failed",
            cleanup=CleanupReport(
                document_resources=document_resources
            ),
            error=(
                None
                if cleanup_succeeded
                else ActionError(
                    code="SESSION_CLEANUP_INCOMPLETE",
                    message=(
                        "Not all Session-owned document resources were "
                        "confirmed released"
                    ),
                )
            ),
        )
        return self._cleanup_outcome

    def close(self) -> CleanupOutcome:
        with self._close_lock:
            return self._close_once()


class ActionSession:
    """Own one Runtime and construct complete caller-facing responses."""

    def __init__(
        self,
        *,
        application: str,
        contracts: ApplicationContractSet,
        adapter: ApplicationAdapter,
        coordinator: Any,
        session_id: str,
        request_id_factory: Optional[Callable[[], str]] = None,
        clock: Optional[Callable[[], float]] = None,
        action_timeout_seconds: float = 120,
        cleanup_timeout_seconds: float = 5,
        launcher: Any = None,
    ):
        self._session_id = session_id
        self._adapter = adapter
        self._launcher = launcher
        self._close_lock = threading.Lock()
        self._cleanup_outcome = None
        self._runtime = _ActionRuntime(
            application=application,
            contracts=contracts,
            adapter=adapter,
            coordinator=coordinator,
            request_id_factory=request_id_factory,
            clock=clock,
            action_timeout_seconds=action_timeout_seconds,
            cleanup_timeout_seconds=cleanup_timeout_seconds,
        )

    def execute(
        self,
        request: ActionRequest,
        trace: TraceContext,
    ) -> ActionTurn:
        runtime_turn = self._runtime.execute(request, trace)
        disposition = runtime_turn.disposition
        response = ActionResponse(
            outcome=disposition.outcome,
            address=disposition.address,
            session_id=self._session_id,
            trace=disposition.trace,
            error=disposition.error,
            data=disposition.data,
        )
        return ActionTurn(
            response=response,
            continuation=runtime_turn.continuation,
        )

    def close(self) -> CleanupOutcome:
        with self._close_lock:
            if self._cleanup_outcome is not None:
                return self._cleanup_outcome
            runtime_cleanup = self._runtime.close()
            try:
                self._adapter.close()
            except Exception:
                pass
            processes = ()
            launcher_failed = False
            if self._launcher is not None:
                try:
                    processes = tuple(self._launcher.close())
                    launcher_failed = any(
                        not process.released for process in processes
                    )
                except Exception:
                    launcher_failed = True
            failed = runtime_cleanup.outcome == "failed" or launcher_failed
            self._cleanup_outcome = CleanupOutcome(
                outcome="failed" if failed else "succeeded",
                cleanup=CleanupReport(
                    document_resources=(
                        runtime_cleanup.cleanup.document_resources
                    ),
                    processes=processes,
                ),
                error=(
                    ActionError(
                        code="SESSION_CLEANUP_INCOMPLETE",
                        message=(
                            "Not all Session-owned automation resources "
                            "were confirmed released"
                        ),
                    )
                    if failed
                    else None
                ),
            )
            return self._cleanup_outcome
