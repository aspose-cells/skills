"""verify_syntax.py

Extracts every ```python code block from the reference docs and compiles it
with the standard library compiler. Catches syntax errors (unbalanced
parens, bad indentation, bad string escapes) without running the code.

Usage: python verify_syntax.py
"""

import pathlib
import re

REFS = pathlib.Path(__file__).parent.parent / "references"


def main():
    bad = 0
    for f in sorted(REFS.glob("*.md")):
        txt = f.read_text(encoding="utf-8")
        blocks = re.findall(r"```python(.*?)```", txt, re.S)
        for i, b in enumerate(blocks):
            src = b.strip()
            if not src:
                continue
            try:
                compile(src, f"<{f.name}:{i}>", "exec")
            except SyntaxError as e:
                bad += 1
                print(f"{f.name} block {i}: {e.msg} line {e.lineno}: {(e.text or '').rstrip()}")
    print("SYNTAX ERRORS:", bad)


if __name__ == "__main__":
    main()
