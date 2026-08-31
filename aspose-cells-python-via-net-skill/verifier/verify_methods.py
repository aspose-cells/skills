"""verify_methods.py

Best-effort method/attribute checker for the Aspose.Cells Python skill.

Unlike runtime introspection (which collapses every .NET overload to
(*args, **kwargs)), this parses the shipped __init__.pyi stubs, builds a
class -> (methods, properties, property-return-types, __getitem__ types)
index, and then resolves the receiver type of every call in the reference
code blocks by following attribute/subscript/method-return chains.

It only reports a call as an error when the receiver type IS resolved to a
known class AND the called name is NOT a member of that class (or its
bases). Unresolvable receivers are skipped, so false positives stay low.
"""

import ast
import pathlib
import sys

ASPOSE_ROOT = pathlib.Path(__import__("aspose.cells", fromlist=["__file__"]).__file__).parent

# class full name -> info
CLASSES = {}


def module_of(pyi_path):
    rel = pyi_path.relative_to(ASPOSE_ROOT.parent)  # e.g. aspose/cells/drawing/__init__.pyi
    parts = list(rel.with_suffix("").parts)
    if parts[-1] == "__init__":
        parts = parts[:-1]
    return ".".join(parts)


def ann_to_str(node):
    if node is None:
        return None
    if isinstance(node, ast.Attribute):
        return ann_to_str(node.value) + "." + node.attr
    if isinstance(node, ast.Name):
        return node.id
    if isinstance(node, ast.Subscript):
        # Container[Inner] -> use the inner type when present
        inner = ann_to_str(node.slice)
        base = ann_to_str(node.value)
        if inner:
            return inner
        return base
    if isinstance(node, ast.Constant) and isinstance(node.value, str):
        return node.value
    return None


def normalize(t):
    if t is None:
        return None
    t = t.strip()
    if "[" in t:
        t = t[: t.index("[")]
    return t


def lookup(t):
    if t is None:
        return None
    t = normalize(t)
    if t in CLASSES:
        return CLASSES[t]
    alt = "aspose.cells." + t
    if alt in CLASSES:
        return CLASSES[alt]
    return None


def parse_pyi(path):
    mod = module_of(path)
    try:
        tree = ast.parse(path.read_text(encoding="utf-8", errors="ignore"))
    except Exception:
        return
    for node in ast.walk(tree):
        if isinstance(node, ast.ClassDef):
            full = mod + "." + node.name
            info = CLASSES.setdefault(
                full,
                {"methods": set(), "members": {}, "bases": [], "getitem": None},
            )
            for base in node.bases:
                b = ann_to_str(base)
                if b:
                    info["bases"].append(normalize(b))
            for item in node.body:
                if isinstance(item, ast.FunctionDef):
                    rt = ann_to_str(item.returns)
                    info["members"][item.name] = rt
                    info["methods"].add(item.name)
                    if item.name == "__getitem__":
                        info["getitem"] = rt
                elif isinstance(item, ast.Assign):
                    for tgt in item.targets:
                        if isinstance(tgt, ast.Name):
                            info["members"].setdefault(tgt.id, None)
            # nested classes (e.g. enums inside a class)
            for sub in node.body:
                if isinstance(sub, ast.ClassDef):
                    sub_full = full + "." + sub.name
                    sinfo = CLASSES.setdefault(
                        sub_full,
                        {"methods": set(), "members": {}, "bases": [], "getitem": None},
                    )
                    for sit in sub.body:
                        if isinstance(sit, ast.Assign):
                            for tgt in sit.targets:
                                if isinstance(tgt, ast.Name):
                                    sinfo["members"].setdefault(tgt.id, None)


def all_members(info):
    """Merge members from this class and all (transitive) bases."""
    seen = {}
    stack = [info]
    visited = set()
    while stack:
        cur = stack.pop()
        if id(cur) in visited:
            continue
        visited.add(id(cur))
        for k, v in cur["members"].items():
            seen.setdefault(k, v)
        for b in cur["bases"]:
            bi = lookup(b)
            if bi:
                stack.append(bi)
    return seen


def build_index():
    for p in ASPOSE_ROOT.rglob("*.pyi"):
        parse_pyi(p)


# ---------------------------------------------------------------------------
# Per-block resolver
# ---------------------------------------------------------------------------

class BlockResolver:
    def __init__(self):
        self.aliases = {}      # alias -> module full name
        self.local_types = {}  # var name -> type string

    def pre_scan(self, tree):
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for n in node.names:
                    if n.asname:
                        self.aliases[n.asname] = n.name
            elif isinstance(node, ast.ImportFrom):
                if node.module and node.names and node.names[0].asname:
                    self.aliases[node.names[0].asname] = node.module

    def resolve_type(self, expr):
        """Return ('type', type_str) or ('module', module_str) or None."""
        if isinstance(expr, ast.Name):
            if expr.id in self.local_types:
                return ("type", self.local_types[expr.id])
            if expr.id in self.aliases:
                return ("module", self.aliases[expr.id])
            return None
        if isinstance(expr, ast.Attribute):
            base = self.resolve_type(expr.value)
            if base is None:
                return None
            if base[0] == "module":
                full = base[1] + "." + expr.attr
                if lookup(full):
                    return ("type", full)
                return ("module", full)
            if base[0] == "type":
                cls = lookup(base[1])
                if cls is None:
                    return None
                members = all_members(cls)
                if expr.attr in members:
                    rt = members[expr.attr]
                    if rt:
                        return ("type", rt)
                    return ("type", base[1] + "." + expr.attr)
                return None
            return None
        if isinstance(expr, ast.Subscript):
            base = self.resolve_type(expr.value)
            if base and base[0] == "type":
                cls = lookup(base[1])
                if cls and cls["getitem"]:
                    return ("type", cls["getitem"])
            return None
        if isinstance(expr, ast.Call):
            f = expr.func
            if isinstance(f, ast.Attribute):
                base = self.resolve_type(f.value)
                if base is None:
                    return None
                if base[0] == "module":
                    # constructor: module.Class(...)
                    full = base[1] + "." + f.attr
                    if lookup(full):
                        return ("type", full)
                    return None
                if base[0] == "type":
                    cls = lookup(base[1])
                    if cls is None:
                        return None
                    members = all_members(cls)
                    if f.attr in members and members[f.attr]:
                        return ("type", members[f.attr])
                    return None
            # plain function call -> unknown
            return None
        return None

    def infer_assignments(self, tree, passes=4):
        for _ in range(passes):
            for node in ast.walk(tree):
                if isinstance(node, ast.Assign):
                    if len(node.targets) != 1 or not isinstance(node.targets[0], ast.Name):
                        continue
                    name = node.targets[0].id
                    t = self.resolve_type(node.value)
                    if t and t[0] == "type":
                        self.local_types[name] = t[1]


def check_block(src, fname, errors):
    try:
        tree = ast.parse(src)
    except SyntaxError:
        return
    res = BlockResolver()
    res.pre_scan(tree)
    res.infer_assignments(tree)
    for node in ast.walk(tree):
        if not isinstance(node, ast.Call):
            continue
        func = node.func
        if not isinstance(func, ast.Attribute):
            continue
        recv = res.resolve_type(func.value)
        if recv is None or recv[0] != "type":
            continue
        cls = lookup(recv[1])
        if cls is None:
            continue
        check_block.resolved += 1
        members = all_members(cls)
        if func.attr not in members:
            # Stub may omit inherited members; fall back to the real class.
            r = runtime_has(recv[1], func.attr)
            if r is False:
                errors.append(
                    f"{fname}: '{func.attr}' is not a member of {recv[1]} "
                    f"(call at line {getattr(node, 'lineno', '?')})"
                )


def runtime_has(full, name):
    """Fallback: does the real class expose this attribute?
    Returns True/False when decidable, None when the module/class can't be loaded."""
    try:
        modname, clsname = full.rsplit(".", 1)
        mod = __import__(modname, fromlist=[clsname])
        cls = getattr(mod, clsname, None)
        if cls is None:
            return None
        return hasattr(cls, name)
    except Exception:
        return None


def main():
    build_index()
    refs = pathlib.Path(__file__).parent.parent / "references"
    errors = []
    check_block.resolved = 0
    for f in sorted(refs.glob("*.md")):
        txt = f.read_text(encoding="utf-8")
        blocks = __import__("re").findall(r"```python(.*?)```", txt, __import__("re").S)
        for b in blocks:
            b = b.strip()
            if b:
                check_block(b, f.name, errors)
    # de-dupe
    errors = sorted(set(errors))
    print(f"classes indexed: {len(CLASSES)}")
    print(f"calls resolved to a known type and checked: {check_block.resolved}")
    print(f"potential method errors: {len(errors)}")
    for e in errors:
        print(" -", e)


if __name__ == "__main__":
    main()
