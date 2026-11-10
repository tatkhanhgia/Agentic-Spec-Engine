#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Convert the backend OpenAPI spec into the document contract shape.

The backend keeps a hand-written OpenAPI 3.x spec (`docs/api/openapi.yaml` in
GoPaperless-Workflow-API), guarded by `OpenApiSpecConsistencyTest` against the
`validURLs` route registry and the DTO snapshot. This script maps it onto the
same record shape `extract_doc_contract.py` emits for the DOCX, so the two can
be compared field by field and, later, rendered:

  * tag            -> category (Heading 2)
  * summary        -> name (Heading 3)
  * header/path    -> headers / path_params tables
  * requestBody    -> request_fields (top-level properties only; nested objects
                      appear as a composite Type, as in the document)
  * 2xx response   -> response_fields / response_headers
  * required       -> presence M/O

Operations without a security requirement override inherit the global one and
get an `Authorization` header row, as the document lists it.

THIS SCRIPT ONLY READS the spec. Output defaults to contract/from_openapi.json.

Usage:
    python scripts/openapi_to_contract.py <openapi.yaml> [-o out.json]
"""
import json
import os
import sys

import yaml

from extract_doc_contract import to_key

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DEFAULT = os.path.join(ROOT, "contract", "from_openapi.json")

HTTP_METHODS = ("get", "post", "put", "delete", "patch")

PRIMITIVE_NAMES = {
    "string": "String",
    "integer": "Integer",
    "boolean": "Boolean",
    "number": "Number",
    "object": "Object",
    "array": "Array",
}


class Spec:
    def __init__(self, data):
        self.data = data

    def resolve(self, node):
        """Follow a local `$ref` chain to the node it points at."""
        seen = set()
        while isinstance(node, dict) and "$ref" in node:
            ref = node["$ref"]
            if ref in seen or not ref.startswith("#/"):
                break
            seen.add(ref)
            target = self.data
            for part in ref[2:].split("/"):
                target = target[part]
            node = target
        return node

    def type_name(self, schema):
        """Render a schema as the document's Type cell text."""
        if not isinstance(schema, dict):
            return ""
        if "$ref" in schema:
            return schema["$ref"].rsplit("/", 1)[-1]
        t = schema.get("type")
        if isinstance(t, list):
            non_null = [x for x in t if x != "null"]
            t = non_null[0] if len(non_null) == 1 else (non_null[0] if non_null else "null")
        if t == "array":
            return "List<%s>" % (self.type_name(schema.get("items") or {}) or "Object")
        # The document writes every integer width as "Integer", so int64 is not
        # rendered as "Long" here.
        return PRIMITIVE_NAMES.get(t, t or "Object")

    def fields(self, schema):
        """Top-level properties of an object schema as attribute rows."""
        schema = self.resolve(schema or {})
        if schema.get("type") == "array" or schema.get("type") == ["array", "null"]:
            return [], self.type_name(schema)
        props = schema.get("properties") or {}
        required = set(schema.get("required") or [])
        rows = []
        for name, prop in props.items():
            resolved = self.resolve(prop)
            rows.append({
                "name": name,
                "type": self.type_name(prop),
                "presence": "M" if name in required else "O",
                "description": " ".join(str(resolved.get("description", "")).split()),
            })
        return rows, ""


def param_row(spec, param):
    p = spec.resolve(param)
    return {
        "name": p.get("name", ""),
        "type": spec.type_name(p.get("schema") or {}),
        "presence": "M" if p.get("required") else "O",
        "description": " ".join(str(p.get("description", "")).split()),
    }, p.get("in")


def json_content(node):
    content = (node or {}).get("content") or {}
    return content.get("application/json") or next(iter(content.values()), None) or {}


def success_response(spec, responses):
    for code in sorted(str(c) for c in responses):
        if code.startswith("2"):
            return code, spec.resolve(responses[code] if code in responses else responses[int(code)])
    return None, {}


def convert(data):
    spec = Spec(data)
    global_security = data.get("security") or []
    tag_order = [t["name"] for t in data.get("tags", [])]
    endpoints = []
    for path, item in data.get("paths", {}).items():
        shared_params = item.get("parameters", [])
        for method in HTTP_METHODS:
            op = item.get(method)
            if not op:
                continue
            headers, path_params = [], []
            security = op.get("security", global_security)
            if security:
                headers.append({
                    "name": "Authorization",
                    "type": "String",
                    "presence": "M",
                    "description": "Bearer access token.",
                })
            for param in shared_params + op.get("parameters", []):
                row, where = param_row(spec, param)
                if where == "header":
                    headers.append(row)
                elif where == "path":
                    path_params.append(row)

            request_fields, request_type = [], ""
            body = spec.resolve(op.get("requestBody") or {})
            if body:
                request_fields, request_type = spec.fields(json_content(body).get("schema"))

            code, resp = success_response(spec, op.get("responses") or {})
            response_fields, response_type = spec.fields(json_content(resp).get("schema"))
            response_headers = []
            for name, h in (resp.get("headers") or {}).items():
                h = spec.resolve(h)
                response_headers.append({
                    "name": name,
                    "type": spec.type_name(h.get("schema") or {}),
                    "presence": "M" if h.get("required") else "O",
                    "description": " ".join(str(h.get("description", "")).split()),
                })

            tags = op.get("tags") or [""]
            endpoints.append({
                "name": op.get("summary", ""),
                "category": tags[0],
                "category_order": tag_order.index(tags[0]) if tags[0] in tag_order else len(tag_order),
                "operation_id": op.get("operationId", ""),
                "description": " ".join(str(op.get("description", "")).split()),
                "method": method.upper(),
                "path": "{baseURL}" + path,
                "key": "%s %s" % (method.upper(), to_key(path)),
                "route": op.get("x-route", ""),
                "headers": headers,
                "path_params": path_params,
                "request_fields": request_fields,
                "request_type": request_type,
                "request_example": json_content(body).get("example") if body else None,
                "success_status": code,
                "response_fields": response_fields,
                "response_type": response_type,
                "response_headers": response_headers,
                "response_example": json_content(resp).get("example"),
                "error_codes": op.get("x-error-codes") or [],
                "needs_review": op.get("x-needs-review") or [],
            })
    return {
        "source": "openapi",
        "version": (data.get("info") or {}).get("version", ""),
        "count": len(endpoints),
        "endpoints": endpoints,
    }


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    if not args:
        print(__doc__)
        sys.exit(2)
    out_path = OUT_DEFAULT
    if "-o" in sys.argv:
        out_path = sys.argv[sys.argv.index("-o") + 1]
        args = [a for a in args if a != out_path]
    with open(args[0], encoding="utf-8") as f:
        data = yaml.safe_load(f)
    contract = convert(data)
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(contract, f, ensure_ascii=False, indent=2)
    print("operations converted : %d" % contract["count"])
    print("wrote %s" % out_path)


if __name__ == "__main__":
    main()
