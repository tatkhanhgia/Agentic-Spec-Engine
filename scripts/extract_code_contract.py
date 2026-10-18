#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Extract the API contract from the GoPaperless Workflow backend source.

The backend is servlet-based (JBoss, not Spring), so there is no Swagger
document to read - the Swagger config in that project is commented out and
only 3 of 25 controllers carry JAX-RS annotations. Routing is instead declared
in two places, and both are easy to read reliably:

  1. `web.xml` maps a servlet to one or more url-pattern prefixes
     (`AdminController` -> `/v1/admin/*`), or the servlet declares them itself
     with `@WebServlet(urlPatterns = {...})`.

  2. Each controller holds a static registry of the exact routes it serves:

         validURLs.put(byteMap[0], ".*/admin/accounts");

     and each `doGet`/`doPost`/`doPut`/`doDelete` tests the incoming URI
     against those entries, naming the operation right after the match:

         if (req.getRequestURI().matches("^" + ... + validURLs.get(byteMap[3]))) {
             loggingObject.apiName = "Lookup Enterprise By Tax Code For Tools";

     A route whose pattern is a literal path (no regex metacharacters) may be
     dispatched through a per-controller helper instead of an inline regex
     match, which reads the registry the same way but hides the `validURLs.get`
     call behind a method boundary:

         if (isExactRoute(req, byteMap[5])) {
             loggingObject.apiName = "Create Personal Template";

     Both forms are real routes and both must be picked up - matching only the
     inline form silently dropped `POST /v1/settings/templates` and
     `POST /v1/enterprise/templates` from the extracted contract.

This registry is the most trustworthy source available: a route that is not
declared here is rejected at runtime, so it cannot drift out of date the way a
hand-exported Postman collection can.

THIS SCRIPT ONLY READS the backend project. It never writes to it, never runs
Maven, and never touches its git state. Output goes to contract/from_code.json
inside this project.

Usage:
    python scripts/extract_code_contract.py [path-to-backend] [-o out.json]
"""
import json
import os
import re
import sys

DEFAULT_BACKEND = (r"C:\Users\Admin\Documents\NetBeansProjects\Library_RSSP_SDK"
                   r"\ProjectRSSP_newest\Gopaperless_Workflow_Maven")

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DEFAULT = os.path.join(ROOT, "contract", "from_code.json")

HTTP_METHODS = ("Get", "Post", "Put", "Delete", "Patch", "Head")

# How a doXxx block names the registry entry it is about to serve. Two forms
# occur: the inline `validURLs.get(byteMap[N])` used by most controllers, and a
# helper such as TemplateController's `isExactRoute(req, byteMap[N])` used for
# literal (non-regex) paths. Any call that passes a byteMap slot counts.
ROUTE_REF = re.compile(r"\w+\s*\([^();]*byteMap\[(\d+)\]")


def read(path):
    with open(path, encoding="utf-8", errors="replace") as fh:
        return fh.read()


def strip_xml_comments(text):
    return re.sub(r"<!--.*?-->", "", text, flags=re.S)


def parse_web_xml(backend):
    """servlet simple class name -> list of url-patterns."""
    path = os.path.join(backend, "src", "main", "webapp", "WEB-INF", "web.xml")
    if not os.path.isfile(path):
        return {}
    xml = strip_xml_comments(read(path))

    name_to_class = dict(re.findall(
        r"<servlet>\s*<servlet-name>([^<]+)</servlet-name>\s*"
        r"<servlet-class>([^<]+)</servlet-class>", xml))

    out = {}
    for name, body in re.findall(
            r"<servlet-mapping>\s*<servlet-name>([^<]+)</servlet-name>(.*?)</servlet-mapping>",
            xml, re.S):
        cls = name_to_class.get(name, name).split(".")[-1]
        out.setdefault(cls, []).extend(
            re.findall(r"<url-pattern>([^<]+)</url-pattern>", body))
    return out


def parse_web_servlet_annotation(src):
    """url-patterns declared by @WebServlet on the class itself."""
    m = re.search(r"@WebServlet\s*\((.*?)\)\s*(?:public\s+)?class", src, re.S)
    if not m:
        return []
    return re.findall(r'"([^"]+)"', m.group(1))


def method_blocks(src):
    """Split a controller into (HTTP method, source slice) pairs."""
    hits = []
    for m in re.finditer(r"\b(?:public|protected)\s+void\s+do(%s)\s*\("
                         % "|".join(HTTP_METHODS), src):
        hits.append((m.group(1).upper(), m.start()))
    blocks = []
    for i, (meth, start) in enumerate(hits):
        end = hits[i + 1][1] if i + 1 < len(hits) else len(src)
        blocks.append((meth, src[start:end], start))
    return blocks


def clean_pattern(pattern):
    """Turn a route regex into a readable path plus a comparable key."""
    p = pattern.strip()
    p = re.sub(r"^\^?\.\*", "", p)          # leading ".*" swallows the context path
    if not p.startswith("/"):
        p = "/" + p
    p = re.sub(r"\$$", "", p)
    return p


def to_key(path):
    """Normalise a path so code and doc compare on equal terms.

    Routes in the registry are regexes, so a path parameter can appear as
    `[0-9]*`, `[^/]+` or a capture group; the document writes the same thing as
    `{packageId}` and Postman fills in a sample value. All three collapse to
    `{p}` here.
    """
    p = path
    m = re.search(r"/v\d+/", p)
    if m:
        p = p[m.start():]
    p = p.replace("(", "").replace(")", "")        # capture groups add no path text
    p = re.sub(r"\{\d+(,\d+)?\}", "", p)           # regex repetition, e.g. {0,1}
    p = re.sub(r"\{\{[^}]*\}\}", "{p}", p)         # {{postmanVar}}
    p = re.sub(r"\[[^\]]*\][*+?]?", "{p}", p)      # [0-9]*, [^/]+, [a-z]+
    p = re.sub(r"\\d[*+?]?", "{p}", p)
    p = re.sub(r"\.[*+]", "{p}", p)
    p = re.sub(r"\{[^}]*\}", "{p}", p)             # {packageId}
    p = re.sub(r":[A-Za-z_][A-Za-z0-9_]*", "{p}", p)
    # A segment spelled as an alternation of literals - `signature|text|checkbox`,
    # `ENTERPRISE|enterprise|PERSONAL|personal` - is a path variable constrained
    # to an enum. The document writes it as one, `{fieldType}` or
    # `{templateType}`, so both sides have to collapse to `{p}` or the endpoint
    # is reported missing when it is in fact documented.
    segs = ["{p}" if (s.isdigit() or "|" in s) else s
            for s in p.split("/") if s != ""]
    return ("/" + "/".join(segs)).rstrip("/").lower()


def version_prefix(url_patterns):
    for up in url_patterns:
        m = re.match(r"(/v\d+)/", up)
        if m:
            return m.group(1)
    return ""


def parse_controller(path, web_map):
    src = read(path)
    cls = os.path.basename(path)[:-5]

    routes = {}
    for idx, pattern in re.findall(
            r"validURLs\.put\(\s*byteMap\[(\d+)\]\s*,\s*\"((?:[^\"\\]|\\.)*)\"", src):
        routes[int(idx)] = pattern
    if not routes:
        return []

    url_patterns = web_map.get(cls) or parse_web_servlet_annotation(src)
    prefix = version_prefix(url_patterns)

    group = os.path.basename(os.path.dirname(path))
    found = []
    for meth, block, block_start in method_blocks(src):
        for m in re.finditer(ROUTE_REF, block):
            idx = int(m.group(1))
            if idx not in routes:
                continue
            after = block[m.end():m.end() + 600]
            name = None
            nm = re.search(r'apiName\s*=\s*"([^"]*)"', after)
            if nm:
                name = nm.group(1)
            if not name:
                before = block[max(0, m.start() - 400):m.start()]
                dm = re.findall(r'desc="([^"]*)"', before)
                if dm:
                    name = dm[-1]

            path_text = clean_pattern(routes[idx])
            if prefix and not re.search(r"/v\d+/", path_text):
                path_text = prefix + path_text
            found.append({
                "method": meth,
                "path": path_text,
                "key": "%s %s" % (meth, to_key(path_text)),
                "name": name or "",
                "controller": cls,
                "group": group,
                "pattern": routes[idx],
            })

    # One route can be matched twice inside the same method (guard clauses);
    # keep the first occurrence of each method+index pair.
    seen = set()
    unique = []
    for e in found:
        k = (e["method"], e["pattern"])
        if k in seen:
            continue
        seen.add(k)
        unique.append(e)
    return unique


def main():
    args = [a for a in sys.argv[1:] if not a.startswith("-")]
    backend = args[0] if args else DEFAULT_BACKEND
    out_path = OUT_DEFAULT
    if "-o" in sys.argv:
        out_path = sys.argv[sys.argv.index("-o") + 1]

    if not os.path.isdir(backend):
        sys.exit("backend not found: %s" % backend)

    web_map = parse_web_xml(backend)
    ctrl_root = os.path.join(backend, "src", "main", "java")

    endpoints = []
    files = 0
    for base, _dirs, names in os.walk(ctrl_root):
        for n in sorted(names):
            if not n.endswith("Controller.java"):
                continue
            full = os.path.join(base, n)
            got = parse_controller(full, web_map)
            if got:
                files += 1
                endpoints.extend(got)

    endpoints.sort(key=lambda e: (e["group"], e["path"], e["method"]))
    data = {
        "source": "code",
        "backend": backend,
        "controllers": files,
        "count": len(endpoints),
        "endpoints": endpoints,
    }
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, ensure_ascii=False, indent=2)

    print("controllers with a route registry : %d" % files)
    print("endpoints extracted               : %d" % len(endpoints))
    unnamed = sum(1 for e in endpoints if not e["name"])
    print("without an apiName                : %d" % unnamed)
    print("wrote %s" % out_path)


if __name__ == "__main__":
    main()
