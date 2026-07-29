#!/usr/bin/env python3
"""
Regenerate Appendix C from the chapters themselves.

Appendix C reprints all ninety affirmations, so transcribing them by hand would
guarantee drift the first time a chapter is edited. Run this after changing any
"Say It Out Loud" section:

    python3 build/sync_affirmations.py
"""

import re
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
CHAPTERS = ROOT / "manuscript" / "chapters"
TARGET = ROOT / "manuscript" / "appendices" / "c-the-affirmations.md"

MOVEMENTS = [
    (1, 10, "I — The Ashes", "For the days when the loss is still fresh."),
    (11, 20, "II — The Forge", "For the days when you are rebuilding the machinery."),
    (21, 30, "III — The Return", "For the days when you are back at the desk."),
]

PREAMBLE = """# Appendix C — The Thirty Affirmations

All ninety, collected. This page is the reason the book does not end on Day 30.

**How to use it.** Read one movement aloud each week and rotate — Ashes, Forge,
Return, and back to the beginning. Say them out loud, in your own voice, where
you can hear yourself. Day 14 explains why the saying rather than the believing
is the mechanism, and why an affirmation your mind can refute will cost you
rather than help you.

Where a line is marked *[Name]*, use your own name. Day 13 explains why the
pronoun is doing real work there.
"""


def main():
    chapters = sorted(CHAPTERS.glob("day-*.md"))
    if len(chapters) != 30:
        sys.exit(f"expected 30 chapters, found {len(chapters)}")

    out = [PREAMBLE]
    for lo, hi, name, blurb in MOVEMENTS:
        out.append(f"\n---\n\n## Movement {name}\n\n{blurb}\n")
        for path in chapters:
            day = int(re.match(r"day-(\d+)", path.name).group(1))
            if not lo <= day <= hi:
                continue
            text = path.read_text(encoding="utf-8")

            title = re.search(r"^# (.+)$", text, re.M).group(1).strip()
            block = re.search(r"## Say It Out Loud\n(.*?)\n## ", text, re.S)
            if not block:
                sys.exit(f"{path.name}: no 'Say It Out Loud' section")
            lines = re.findall(r"^\d+\.\s+(.*)$", block.group(1), re.M)
            if len(lines) != 3:
                sys.exit(f"{path.name}: expected 3 affirmations, found {len(lines)}")

            out.append(f"\n### Day {day} — {title}\n")
            out.extend(f"{i}. {t}" for i, t in enumerate(lines, 1))
            out.append("")

    TARGET.write_text("\n".join(out).rstrip() + "\n", encoding="utf-8")
    print(f"  wrote {TARGET.relative_to(ROOT)}  (90 affirmations from 30 chapters)")


if __name__ == "__main__":
    main()
