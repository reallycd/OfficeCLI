# Paragraph Formatting Showcase

Exercises the docx **paragraph** property surface. Three files:

- **paragraph-showcase.sh** — builds the document with `officecli`.
- **paragraph-showcase.docx** — generated output.
- **paragraph-showcase.md** — this file.

## Regenerate

```bash
cd examples/word
bash paragraph-showcase.sh
# → paragraph-showcase.docx
```

## Sections

| Section | Properties |
|---|---|
| Alignment | `align=left\|center\|right\|both` |
| Indentation | `indent` (left), `rightIndent`, `firstLineIndent`, `hangingIndent` |
| Spacing | `spaceBefore`, `spaceAfter`, `lineSpacing` |
| Pagination flags | `keepNext`, `keepLines`, `widowControl` |
| Paragraph-level run formatting | `bold`, `italic`, `color`, `size`, `highlight` (applied to every run) |
| Shading | `shading.fill` |
| Paragraph-mark formatting | `markRPr.bold`, `markRPr.color` (the ¶ glyph only) |
| Outline level | `outlineLvl` |

## Two kinds of "bold" on a paragraph

- **`bold`** applies to every run in the paragraph (and reads back as `bold`).
- **`markRPr.bold`** formats only the paragraph mark (the ¶ pilcrow) — distinct
  from run text, used so appended runs inherit the mark's formatting. They are
  independent: setting one does not surface as the other.

```bash
officecli set file.docx /body/p[1] --prop bold=true          # all runs bold
officecli set file.docx /body/p[1] --prop markRPr.bold=true   # ¶ mark only
```

> `shading.fill` alone now produces schema-valid output — the writer defaults
> the required `w:shd/@val` to `clear`. Pair with `shading.val` for a pattern
> shade.
