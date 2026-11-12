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
            target = self.resolve(schema)
            if self.is_enum(target):
                # A named enumeration is a primitive to the reader; the field's
                # `enum` entry carries the name.
                return PRIMITIVE_NAMES.get(target.get("type"), "String")
            return target.get("x-doc-type") or schema["$ref"].rsplit("/", 1)[-1]
        t = schema.get("type")
        if isinstance(t, list):
            non_null = [x for x in t if x != "null"]
            t = non_null[0] if len(non_null) == 1 else (non_null[0] if non_null else "null")
        if t == "array":
            return "List<%s>" % (self.type_name(schema.get("items") or {}) or "Object")
        # The document writes every integer width as "Integer", so int64 is not
        # rendered as "Long" here.
        return PRIMITIVE_NAMES.get(t, t or "Object")

    @staticmethod
    def is_enum(schema):
        """A named enumeration: a component schema with values and no properties."""
        return isinstance(schema, dict) and "enum" in schema and "properties" not in schema

    def ref_name(self, schema):
        """Object schema a property points at (directly or as array items), or None."""
        if not isinstance(schema, dict):
            return None
        target = schema.get("items") if "items" in schema else schema
        ref = (target or {}).get("$ref", "")
        if ref.startswith("#/components/schemas/") and not self.is_enum(self.resolve(target)):
            return ref.rsplit("/", 1)[-1]
        return None

    def enum_name(self, schema):
        """Named enumeration a property takes its values from, or None.

        The reference is the property itself, its array items, or an `allOf`
        member (the form that keeps `type` beside the reference).
        """
        if not isinstance(schema, dict):
            return None
        target = schema.get("items") if "items" in schema else schema
        for node in [target] + list((target or {}).get("allOf") or []):
            ref = (node or {}).get("$ref", "") if isinstance(node, dict) else ""
            if ref.startswith("#/components/schemas/") and self.is_enum(self.resolve(node)):
                return ref.rsplit("/", 1)[-1]
        return None

    def collect(self, schema, props, required, conditional, seen=None):
        """Merge properties and required names of a schema and its allOf chain.

        A name required only in some anyOf/oneOf branch is conditional: the
        document shows it as RC rather than M or O.
        """
        seen = seen if seen is not None else set()
        ref = schema.get("$ref") if isinstance(schema, dict) else None
        if ref:
            if ref in seen:
                return
            seen.add(ref)
        schema = self.resolve(schema or {})
        for part in schema.get("allOf") or []:
            self.collect(part, props, required, conditional, seen)
        for name, prop in (schema.get("properties") or {}).items():
            props.setdefault(name, prop)
        required.update(schema.get("required") or [])
        for key in ("anyOf", "oneOf"):
            for branch in schema.get(key) or []:
                b_props, b_required, b_conditional = {}, set(), set()
                self.collect(branch, b_props, b_required, b_conditional, set(seen))
                conditional.update(b_required | b_conditional)

    def fields(self, schema):
        """Top-level properties of an object schema as attribute rows."""
        schema = schema or {}
        resolved = self.resolve(schema)
        if resolved.get("type") == "array" or resolved.get("type") == ["array", "null"]:
            return [], self.type_name(resolved)
        props, required, conditional = {}, set(), set()
        self.collect(schema, props, required, conditional)
        rows = []
        for name, prop in props.items():
            resolved = self.resolve(prop)
            row = {
                "name": name,
                "type": self.type_name(prop),
                "presence": ("M" if name in required
                             else "RC" if name in conditional
                             else "O"),
                # A description beside a $ref describes this field; the target's
                # own description describes the object.
                "description": " ".join(str(prop.get("description")
                                            or resolved.get("description", "")).split()),
            }
            ref = self.ref_name(prop)
            if ref:
                row["ref"] = ref
            enum = self.enum_name(prop)
            if enum:
                row["enum"] = enum
            example = prop.get("example", resolved.get("example"))
            if example is not None:
                row["example"] = example
            rows.append(row)
        return rows, ""


def param_row(spec, param):
    p = spec.resolve(param)
    row = {
        "name": p.get("name", ""),
        "type": spec.type_name(p.get("schema") or {}),
        "presence": "M" if p.get("required") else "O",
        "description": " ".join(str(p.get("description", "")).split()),
    }
    enum = spec.enum_name(p.get("schema") or {})
    if enum:
        row["enum"] = enum
    return row, p.get("in")


def json_content(node):
    content = (node or {}).get("content") or {}
    return content.get("application/json") or next(iter(content.values()), None) or {}


def success_response(spec, responses):
    for code in sorted(str(c) for c in responses):
        if code.startswith("2"):
            return code, spec.resolve(responses[code] if code in responses else responses[int(code)])
    return None, {}


def expand_sections(spec, base, sections):
    """One record per document section when an operation spans several.

    `x-doc-sections` lists `{title, path-values?, when?, request-schema?}`.
    `path-values` fills path parameters, so `/fields/{fieldType}` becomes
    `/fields/text`; `when` names the request variant (e.g. a grant_type) the
    section documents; `request-schema` names the anyOf branch whose fields
    the section's Request table lists.
    """
    if not sections:
        return [base]
    records = []
    for section in sections:
        record = dict(base)
        record["name"] = section.get("title", base["name"])
        path = base["path"]
        values = section.get("path-values") or {}
        for param, value in values.items():
            path = path.replace("{%s}" % param, str(value))
        record["path"] = path
        record["key"] = "%s %s" % (base["method"], to_key(path))
        record["path_params"] = [p for p in base["path_params"] if p["name"] not in values]
        record["variant"] = section.get("when", "")
        if section.get("request-schema"):
            record["request_fields"], record["request_type"] = spec.fields(
                {"$ref": section["request-schema"]})
        records.append(record)
    return records


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
            body = spec.resolve(op.get("requestBody") or {})
            for param in shared_params + op.get("parameters", []):
                row, where = param_row(spec, param)
                # OpenAPI ignores a header parameter named Content-Type; the
                # request body's media types are its source of truth.
                if where == "header" and row["name"].lower() != "content-type":
                    headers.append(row)
                elif where == "path":
                    path_params.append(row)
            if body.get("content"):
                media = list(body["content"])
                headers.append({
                    "name": "Content-Type",
                    "type": "String",
                    "presence": "M" if body.get("required") else "O",
                    "description": "The MIME type of the request body: %s." % " or ".join(media),
                })

            request_fields, request_type = [], ""
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
            base = {
                "name": op.get("x-doc-title") or op.get("summary", ""),
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
                "status": op.get("x-doc-status", ""),
                "media_types": list((body.get("content") or {}).keys()),
                "response_media_types": list((resp.get("content") or {}).keys()),
            }
            endpoints.extend(expand_sections(spec, base, op.get("x-doc-sections")))
    # Object definitions for the RELATIVE OBJECT chapter: every component
    # schema with properties, under the label its Type cells show.
    objects = {}
    for name, schema in ((data.get("components") or {}).get("schemas") or {}).items():
        rows, _ = spec.fields({"$ref": "#/components/schemas/%s" % name})
        if rows and not Spec.is_enum(schema):
            objects[name] = {"label": schema.get("x-doc-type") or name, "fields": rows}
    # Named enumerations for the ENUMERATION chapter: values in declaration
    # order, each with its meaning from `x-enumDescriptions`.
    enums = {}
    for name, schema in ((data.get("components") or {}).get("schemas") or {}).items():
        if Spec.is_enum(schema):
            described = schema.get("x-enumDescriptions") or {}
            enums[name] = {
                "label": schema.get("x-doc-type") or name,
                "description": " ".join(str(schema.get("description", "")).split()),
                "values": [{"value": str(v), "description": " ".join(str(described.get(v, "")).split())}
                           for v in schema["enum"]],
            }
    return {
        "source": "openapi",
        "version": (data.get("info") or {}).get("version", ""),
        "count": len(endpoints),
        "endpoints": endpoints,
        "objects": objects,
        "enums": enums,
    }


def main():
    argv = sys.argv[1:]
    out_path = OUT_DEFAULT
    if "-o" in argv:
        i = argv.index("-o")
        out_path = argv[i + 1]
        argv = argv[:i] + argv[i + 2:]
    args = [a for a in argv if not a.startswith("-")]
    if not args:
        print(__doc__)
        sys.exit(2)
    with open(args[0], encoding="utf-8") as f:
        data = yaml.safe_load(f)
    contract = convert(data)
    if os.path.exists(out_path) and os.path.samefile(args[0], out_path):
        sys.exit("Output path must not refer to the input spec")
    os.makedirs(os.path.dirname(os.path.abspath(out_path)), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(contract, f, ensure_ascii=False, indent=2)
    print("operations converted : %d" % contract["count"])
    print("wrote %s" % out_path)


if __name__ == "__main__":
    main()
