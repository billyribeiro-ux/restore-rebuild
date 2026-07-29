# Restore & Rebuild

*30 Days from the Blow-Up to the Comeback* — a thirty-day book for traders rebuilding after a blown account.

One short chapter a day. Each one carries a verse, one mechanism of the brain explained in plain language, a scripted meditation, three spoken affirmations, and one small action. Three movements of ten days: **The Ashes** (grief and release), **The Forge** (retraining the mind), **The Return** (rebuilding trust and coming back slowly).

The practical trading protocol is deliberately kept out of the chapters and lives in Appendix A, so the daily reading stays inner work.

## Layout

```
manuscript/
  order.txt        reading order — the single source of truth for sequence
  front/           title page, opening letter, how to use, the necessary word
  chapters/        part dividers and the thirty daily chapters
  appendices/      re-entry protocol, journal templates, affirmations, indexes
build/
  build.py         Markdown -> styled DOCX, plus the concatenated Markdown
  build.sh         runs build.py, then converts to PDF
  fonts.conf       font substitution so the PDF matches the DOCX design
dist/              the built book (committed, so it can be read without a build)
outline.md         the thirty-day map
```

## Building

```bash
bash build/build.sh
```

Produces `dist/restore-and-rebuild.docx`, `.pdf`, and `.md`.

Requires Python 3 with `python-docx` (installed automatically by the script) and LibreOffice Writer for the PDF step:

```bash
apt-get install -y --no-install-recommends libreoffice-writer
```

Without Writer the DOCX and Markdown still build; only the PDF step fails.

The build prints a word count for every chapter and flags any outside the 750–1,100 word band, which is the length that keeps a chapter to a four-minute read.

## Revising

- **Reordering or adding chapters** — edit `manuscript/order.txt`. Filenames never need to change.
- **Chapter structure** — every daily chapter uses the same eight `##` sections. Three of them are styled specially by name: `The Meditation`, `Say It Out Loud`, and `Carry This`. If you rename those headings, update the `SEC_*` constants at the top of `build/build.py`.
- **Page design** — all typography lives in `build_styles()` in `build/build.py`. Page size is 5.5 × 8.5 in, the standard digest trim for a devotional.
- **Author name** — replace `[ Author Name ]` in `manuscript/front/00-title.md`.

## Scripture

Verses are from the **World English Bible (WEB)**, which is in the public domain and can be quoted without permission or licence fees. This keeps the book publishable as it stands. Modern copyrighted translations (NIV, ESV, NLT) have quoting limits that would need clearing with the publisher first; every verse sits in its own `>` block, so swapping translations later is a mechanical edit.

## A note on the research

Every psychological claim in the book is tied to named published work, listed chapter by chapter in Appendix E. Where the evidence is contested — grit, ego depletion, the marshmallow test — the chapter says so rather than presenting a tidier story than the literature supports.
