# SPDX-License-Identifier: Apache-2.0
"""Validate the JSON Schema vocabulary used by the preserved Office tool schemas.

Reused unchanged in behaviour from hackathon-office-live@ed09456 validation.py;
only the error class import differs. Office.js remains the authority for
host-specific rules; defaults are never inserted into user arguments.
"""
import math

from .bridge_server import BridgeError


def validate(value, schema, path="arguments"):
    def fail():
        raise BridgeError("Validation", f"{path} does not satisfy the canonical tool schema")

    t = schema.get("type")
    checks = {
        "object": lambda: isinstance(value, dict),
        "array": lambda: isinstance(value, list),
        "string": lambda: isinstance(value, str),
        "boolean": lambda: isinstance(value, bool),
        "integer": lambda: isinstance(value, int) and not isinstance(value, bool),
        "number": lambda: isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value),
    }
    if t is not None and (t not in checks or not checks[t]()):
        fail()
    if "enum" in schema and value not in schema["enum"]:
        fail()
    if isinstance(value, dict):
        if any(k not in value for k in schema.get("required", [])):
            fail()
        if len(value) < schema.get("minProperties", 0):
            fail()
        properties = schema.get("properties", {})
        extra = schema.get("additionalProperties", True)
        for key, v in value.items():
            if not isinstance(key, str):
                fail()
            if key in properties:
                validate(v, properties[key], path + "." + key)
            elif extra is False:
                fail()
            elif isinstance(extra, dict):
                validate(v, extra, path + "." + key)
    if isinstance(value, list):
        if len(value) < schema.get("minItems", 0) or len(value) > schema.get("maxItems", float("inf")):
            fail()
        for i, v in enumerate(value):
            if isinstance(schema.get("items"), dict):
                validate(v, schema["items"], f"{path}[{i}]")
    if isinstance(value, str):
        if len(value) < schema.get("minLength", 0) or len(value) > schema.get("maxLength", float("inf")):
            fail()
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if not math.isfinite(value) or value < schema.get("minimum", -float("inf")) \
                or value > schema.get("maximum", float("inf")):
            fail()
