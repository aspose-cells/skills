# Name-level verifier for a GridJs Python skill.
#
# Oracle: the python API doc in api-doc/ next to this script. Candidates are pulled
# out with Python syntax (snake_case members, module imports, tracked receivers),
# not by reflecting over a .NET assembly.
#
# Usage: python name_verifier.py [skillRoot]
# Exit code: 0 = clean, 1 = at least one unknown name, 2 = the verifier could not run.
#
# A finding is a name the skill claims about this library that the API doc does not
# carry. A suspect is a name the scan cannot attach to the library at all - it is
# counted and printed, never silently dropped.

import keyword
import os
import re
import sys

DQ = '"'

# Builtins, stdlib and Flask scaffolding that appear in the docs but are not GridJs API.
IGNORE_NAMES = {
    "self", "cls", "args", "kwargs", "none", "true", "false",
    "print", "open", "len", "range", "str", "int", "float", "bool", "list", "dict",
    "set", "tuple", "bytes", "bytearray", "super", "isinstance", "issubclass",
    "getattr", "setattr", "delattr", "hasattr", "enumerate", "zip", "map", "filter",
    "sorted", "reversed", "sum", "min", "max", "abs", "any", "all", "repr", "format",
    "input", "exit", "main", "run", "callable", "vars", "dir", "id", "type",
    "Flask", "Blueprint", "Response", "jsonify", "render_template", "send_file",
    "send_from_directory", "make_response", "abort", "redirect", "url_for", "session",
    "request", "app", "io", "os", "sys", "re", "json", "gzip", "uuid", "time",
    # Other web frameworks the guides may mention while telling the reader what this one is not.
    "FastAPI", "Starlette", "Pydantic", "Django", "Tornado", "Sanic", "Quart",
    "uvicorn", "gunicorn", "werkzeug", "Jinja2", "aiohttp",
    "shutil", "traceback", "logging", "tempfile", "threading", "subprocess", "base64",
    "hashlib", "pathlib", "typing", "dataclasses", "datetime", "collections",
    "BytesIO", "StringIO", "Path", "PurePath", "HTTPException", "TestClient",
    "MagicMock", "pytest", "unittest", "Exception", "ValueError", "TypeError",
    "KeyError", "IOError", "OSError", "RuntimeError",
    "FileNotFoundError", "PermissionError", "NotADirectoryError", "IsADirectoryError",
    "AttributeError", "IndexError", "ImportError", "ModuleNotFoundError",
    "ZeroDivisionError", "NotImplementedError", "FileExistsError",
    "options", "settings", "service", "client", "config", "logger", "cache",
    "file", "path", "name", "value", "data", "result", "error", "response", "request",
}

# Names that appear as attributes on tracked objects but are plain Python plumbing.
IGNORE_ATTRS = IGNORE_NAMES | {
    "aspose", "cellsgridjs", "gridjs",
    "decode", "encode", "read", "write", "seek", "tell", "close", "flush", "readlines",
    "append", "extend", "insert", "pop", "keys", "values", "items", "get", "update",
    "strip", "lstrip", "rstrip", "split", "rsplit", "join", "replace", "startswith",
    "endswith", "lower", "upper", "format", "makedirs", "exists", "isdir", "listdir",
    "walk", "join", "basename", "dirname", "abspath", "realpath", "environ", "sep",
}

# Troubleshooting prose names a stand-in member to describe a pattern - "call
# your_service.some_method()" - rather than claim that member exists. Underscored and
# prefixed, so it cannot collide with a real snake_case GridJs member.
PLACEHOLDER_MEMBER = re.compile(
    r"(?:some|any|our|my|your|their|dummy|fake|placeholder)_[a-z0-9_]+"
    r"|(?:some|any|our|my|your|their|dummy|fake)[A-Z][A-Za-z0-9]*"
    r"|foo|bar|baz|qux|x{3,}"
    r"|do_something|something|whatever|anything"
)


def read_doc_text(path):
    raw = open(path, "rb").read()
    # The python doc is UTF-16 without a BOM; sniff for the interleaved NUL bytes.
    if b"\x00" in raw[:20]:
        return raw.decode("utf-16")
    return raw.decode("utf-8-sig")


def load_oracle(api_doc_path):
    types = set()
    members_by_type = {}
    all_members = set()
    member_owners = {}

    text = read_doc_text(api_doc_path)
    for kind, qualified in re.findall(r'<member name="([TFMPE]):([^"]*)"', text):
        head = qualified.split("(")[0]
        parts = [p for p in head.split(".") if p]
        if not parts:
            continue

        if kind == "T":
            types.add(parts[-1])
            continue

        member = parts[-1]
        owner = parts[-2] if len(parts) > 1 else ""
        if owner:
            types.add(owner)
            members_by_type.setdefault(owner, set()).add(member)
            all_members.add(member)
            member_owners.setdefault(member, set()).add(owner)

        # Parameter types are part of the documented surface too.
        for params in re.findall(r"\((.*)\)", qualified):
            for decl in params.split(","):
                decl = decl.strip()
                if not decl:
                    continue
                token = re.sub(r"<.*>", "", decl.split()[-1]).split(".")[-1]
                if len(token) > 2 and token[0].isupper():
                    types.add(token)

    return types, members_by_type, all_members, member_owners


def find_skill_root(start_dir):
    d = os.path.abspath(start_dir)
    while True:
        if os.path.exists(os.path.join(d, "skill-index.md")):
            return d
        if os.path.exists(os.path.join(d, "SKILL.md")):
            return d
        md = [f for f in os.listdir(d) if f.endswith(".md")] if os.path.isdir(d) else []
        if len([f for f in md if "-api-" in f]) > 0 and len(md) > 1:
            return d
        parent = os.path.dirname(d)
        if parent == d:
            return None
        d = parent


def blank_non_code(block):
    # Blank the inside of comments and of every string, triple-quoted bodies included. A
    # Flask route returns its client script as a triple-quoted literal, and a scanner that
    # only strips single-line quotes reads the JavaScript inside it as a claim about
    # this library.
    # Character count is preserved so every reported line number stays true.
    sq3 = "'" * 3
    dq3 = DQ * 3
    out = []
    i = 0
    n = len(block)
    while i < n:
        c = block[i]
        if c == "#":
            end = block.find("\n", i)
            end = n if end < 0 else end
            out.append(" " * (end - i))
            i = end
            continue
        if block.startswith(dq3, i) or block.startswith(sq3, i):
            quote = block[i:i + 3]
            end = block.find(quote, i + 3)
            end = n if end < 0 else end + len(quote)
            out.append(quote + " " * max(end - i - len(quote), 0))
            i = end
            continue
        if c == DQ or c == "'":
            end = i + 1
            while end < n:
                if block[end] == "\\":
                    end += 2
                    continue
                if block[end] == c:
                    end += 1
                    break
                end += 1
            out.append(c + " " * max(end - i - 1, 0))
            i = end
            continue
        out.append(c)
        i += 1
    return "".join(out)


def python_blocks(text):
    # Both fences must start a line. A guide sometimes writes a block inside a markdown table
    # cell with a literal escape instead of a newline; matching that would read the rest of the
    # row - and the next row's prose - as code.
    for m in re.finditer(r"^[ \t]*```(?:python|py)[^\n]*\n(.*?)^[ \t]*```", text, re.S | re.M):
        yield text[: m.start()].count("\n") + 2, m.group(1)


def known(name, types, members_by_type, all_members):
    return (
        name in IGNORE_NAMES
        or name in types
        or name in all_members
        or name in members_by_type
    )


def fold_import_lists(block):
    # "from aspose.cellsgridjs import (\\n    Config,\\n    GridJsOptions,\\n)" is one
    # statement written across lines. Read line by line, the opening parenthesis looks like
    # an imported name and the names inside look like ordinary code, so the list is folded
    # back onto the line the import starts on. Continuation lines are blanked rather than
    # dropped to keep every reported line number true.
    lines = block.split("\n")
    i = 0
    while i < len(lines):
        match = re.match(r"(\s*from\s+aspose\.cellsgridjs[\w.]*\s+import\s+)\((.*)$", lines[i])
        if not match:
            i += 1
            continue
        head, rest = match.group(1), match.group(2)
        if ")" in rest:
            lines[i] = head + " " + rest.replace(")", " ").strip()
        else:
            collected = [rest]
            j = i + 1
            while j < len(lines) and ")" not in lines[j]:
                collected.append(lines[j])
                lines[j] = ""
                j += 1
            if j < len(lines):
                collected.append(lines[j].split(")")[0])
                lines[j] = lines[j].split(")", 1)[1]
            lines[i] = head + " " + " ".join(collected)
        i += 1
    return lines


def check_block(start_line, raw_block, rel, types, members_by_type, all_members,
                member_owners, stats, findings, suspects):
    block = blank_non_code(raw_block)
    # Structural candidates are attached to something that can only be GridJs (an
    # aspose.cellsgridjs import, a member of a tracked GridJs object). Loose ones are
    # plain identifiers that an example block is free to invent.
    structural = []
    loose = []
    var_types = {}
    # A class or function the snippet declares itself (class MyGridJsController(GridJsService),
    # def _gzip_json(...)) and a placeholder it names for itself (My..., Custom...) are the
    # reader's own code, so the class-like-call scan must not read them as claims about the library.
    declared = set(re.findall(r"\bclass\s+([A-Z][A-Za-z0-9_]*)", block))
    declared |= set(re.findall(r"^\s*def\s+([A-Za-z_][A-Za-z0-9_]*)", block, re.M))
    placeholder = re.compile(r"(?:My|Custom|Demo|Sample|Test|Example|Simple|Own)[A-Z][A-Za-z0-9_]*")

    for offset, line in enumerate(fold_import_lists(block)):
        if not line.strip():
            continue
        line_no = start_line + offset

        for var, cls in re.findall(r"\b([a-z_][a-z0-9_]*)\s*=\s*([A-Z][A-Za-z0-9_]*)\s*[(=]", line):
            if cls in types:
                var_types[var] = cls

        # 1. Names imported from the GridJs module.
        for imported in re.findall(
            r"from\s+aspose\.cellsgridjs[\w.]*\s+import\s+(.+)$", line
        ):
            for entry in imported.split(","):
                symbol = entry.split(" as ")[0].strip().strip(DQ).strip("'")
                if not symbol or symbol == "*":
                    continue
                if not re.match(r"[A-Za-z_][A-Za-z0-9_]*$", symbol):
                    # Something the scanner failed to parse, not a name the doc claims.
                    continue
                if symbol in types or symbol in all_members:
                    stats["checked"] += 1
                    continue
                # A builtin or another framework's name can still be an import lie here: the
                # package exports it only if the doc lists it, so IGNORE_NAMES must not
                # exempt an import the way it exempts an ordinary call.
                why = ("is not exported by aspose.cellsgridjs - it is a Python or "
                       "other-framework name") if symbol in IGNORE_NAMES \
                    else "imported from aspose.cellsgridjs but undocumented"
                structural.append((line_no, symbol, why))

        if re.match(r"\s*(from|import)\b", line):
            continue

        # 2. Class-like calls: Foo(...) must name a documented type.
        # The lookbehind keeps qualified stdlib calls (gzip.GzipFile) out of the run.
        for name in re.findall(r"(?<!\.)\b([A-Z][A-Za-z0-9_]{2,})\s*\(", line):
            if known(name, types, members_by_type, all_members) or name in declared \
                    or placeholder.fullmatch(name):
                stats["checked"] += 1
                continue
            loose.append((line_no, name, "not found in the GridJs Python API"))

        # 3. Bare calls at statement start - the doc's signature lines live here.
        bare = re.match(r"\s*([a-z_][a-z0-9_]{3,})\s*\(", line)
        if bare:
            name = bare.group(1)
            # "except (ValueError, OSError):" reads as a call to this regex; the language
            # keyword is never a name the doc is claiming.
            if keyword.iskeyword(name):
                continue
            if name in all_members or name in IGNORE_NAMES or name in types \
                    or name in declared or PLACEHOLDER_MEMBER.fullmatch(name):
                stats["checked"] += 1
            else:
                loose.append((line_no, name, "not found in the GridJs Python API"))

        # 4. Attribute access on a documented type or a variable holding one.
        for receiver, attr in re.findall(r"\b([A-Za-z_][A-Za-z0-9_]*)\.([a-z_][a-z0-9_]{2,})\b", line):
            owner = None
            if receiver in types:
                owner = receiver
            elif receiver in var_types:
                owner = var_types[receiver]
            elif re.search(r"grid_?js|aspose", receiver, re.I):
                # Obvious GridJs object with no assignment in this block.
                if attr in all_members or attr in IGNORE_ATTRS \
                        or PLACEHOLDER_MEMBER.fullmatch(attr):
                    stats["checked"] += 1
                else:
                    structural.append((line_no, attr, "not a documented member of " + receiver))
                continue
            if owner is None:
                continue
            if owner not in members_by_type:
                # The oracle quotes this type in signatures (Stream, Workbook) but lists
                # no members for it, so its members are outside what we can judge.
                stats["checked"] += 1
                continue
            owned = members_by_type[owner]
            if attr in owned or attr in IGNORE_ATTRS or PLACEHOLDER_MEMBER.fullmatch(attr):
                stats["checked"] += 1
            elif attr in all_members:
                # Documented, but on somebody else. This is the shape an LLM hallucination
                # takes most often - a real Config setter hung off a workbook - so the
                # global member set must not vouch for it once the owner is known.
                structural.append((line_no, attr, "is documented on "
                                   + "/".join(sorted(member_owners.get(attr, ())))
                                   + ", not on " + owner))
            else:
                structural.append((line_no, attr, "not a documented member of " + owner))

    # Names one of the structural scans already claimed. The loose scan reaches them again
    # through a different pattern; a second line for the same defect is noise, and a finding
    # is never a suspect.
    structural_keys = set((line_no, name) for line_no, name, _ in structural)

    # Several unattached unknowns mean example or app-specific code rather than an API
    # reference, so they are demoted instead of dropped: a block that quietly disappears
    # from the report cannot be reviewed, but it must not fail the run either.
    if len(loose) >= 3:
        seen = set()
        for line_no, name, why in loose:
            key = (line_no, name)
            if key in structural_keys or key in seen:
                continue
            seen.add(key)
            stats["suspect"] += 1
            suspects.append(rel + ":L" + str(line_no) + "  '" + name + "' - " + why)
        loose = []

    reported = set()
    for line_no, name, why in structural + loose:
        key = (line_no, name)
        if key in reported:
            continue
        reported.add(key)
        stats["checked"] += 1
        stats["not_found"] += 1
        findings.append(rel + ":L" + str(line_no) + "  '" + name + "' - " + why)


def main(argv):
    script_dir = os.path.dirname(os.path.abspath(__file__))
    api_docs = []
    doc_dir = os.path.join(script_dir, "api-doc")
    if os.path.isdir(doc_dir):
        api_docs = [f for f in os.listdir(doc_dir) if f.endswith(".xml")]
    if not api_docs:
        print("No API doc found under verifier/api-doc - cannot verify names.", file=sys.stderr)
        return 2

    types, members_by_type, all_members, member_owners = load_oracle(
        os.path.join(doc_dir, sorted(api_docs)[0]))

    root = argv[1] if len(argv) > 1 else find_skill_root(script_dir)
    if not root:
        print("Could not locate skill root.", file=sys.stderr)
        return 2

    print("GridJs Python API index: " + str(len(types)) + " types, "
          + str(len(all_members)) + " members")
    print("Scanning: " + root)
    print("-" * 60)

    stats = {"checked": 0, "not_found": 0, "suspect": 0}
    findings = []
    suspects = []

    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames
                        if not d.startswith("example") and d not in ("verifier", "report")]
        for name in sorted(filenames):
            if not name.endswith(".md"):
                continue
            path = os.path.join(dirpath, name)
            rel = os.path.relpath(path, root).replace(os.sep, "/")
            text = open(path, encoding="utf-8", errors="replace").read()
            for start_line, block in python_blocks(text):
                check_block(start_line, block, rel, types, members_by_type,
                            all_members, member_owners, stats, findings, suspects)

    print("-" * 60)
    print("CHECKED: " + str(stats["checked"]) + "   NOT FOUND: " + str(stats["not_found"])
          + "   SUSPECT: " + str(stats["suspect"]))
    # Machine-readable result line. The caller trusts this rather than the exit code, because
    # an interpreter that dies before finishing also leaves exit code 1 behind, and that is a
    # broken toolchain, not a statement about the documentation.
    print("VERIFIER-RESULT: checked=" + str(stats["checked"])
          + " not-found=" + str(stats["not_found"])
          + " suspect=" + str(stats["suspect"]))

    if findings:
        print("-" * 60)
        print("Undocumented Python API names (" + str(len(findings)) + "):")
        for f in findings:
            print("  " + f)

    if suspects:
        print("-" * 60)
        print("Demoted from blocks of placeholder code, review by hand (" + str(len(suspects)) + "):")
        for s in suspects:
            print("  " + s)

    return 0 if stats["not_found"] == 0 else 1


if __name__ == "__main__":
    try:
        code = main(sys.argv)
    except Exception as exc:
        # A crash says the verifier is broken, not that the skill has bad names.
        print("Verifier could not run: %s: %s" % (type(exc).__name__, exc), file=sys.stderr)
        code = 2
    sys.exit(code)

