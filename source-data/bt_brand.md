# BT Brand Colors

Source: `BT Brand Colors.docx` (Google Drive id 1SRd65l6ZYu04LcLYbMNbjcAPs_YvSCz1)

The document body literally lists three hex codes (no role labels are written in the doc). Roles below are inferred from the colors themselves — treat the hex values as authoritative, the role names as a suggestion.

## Palette (verbatim hex from the document)

| Hex | Appearance | Suggested role |
|-----|-----------|----------------|
| `#2c303b` | Dark charcoal / slate | Primary / dark base (text, sidebar, headers) |
| `#dff5f2` | Pale mint / teal tint | Secondary / light background tint |
| `#d1ec51` | Lime / yellow-green | Accent / highlight |

Extra: `#A6A6A6` (mid grey) appears only as a table shading fill in the file's formatting, not listed as a brand color — ignore unless a neutral grey is needed.

## Fonts

The document specifies NO custom brand font. It uses the Word default theme:
- Headings / major: **Calibri Light**
- Body / minor: **Calibri**

There are no explicit "heading font" / "body font" brand directives in this file. If the dashboard needs a font decision, this docx does not mandate one.

## Note vs. the dashboard's own palette
The existing dashboard CSS (`:root`) uses a different, richer green-based palette (`--bt-green:#00A376`, `--bt-slate:#2C3136`, etc. — see html_outline.md). `#2c303b` here is essentially the dashboard's `--bt-slate`/`--bt-slate-soft` family. The three brand hexes here (charcoal, mint, lime) are the "official" BT brand chips; the dashboard palette is a broader working system.
