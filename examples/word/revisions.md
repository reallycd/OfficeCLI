# revisions

End-to-end demo of the docx revision (track-changes) API, covering **every** revision element the docx handler supports on this branch. Run [revisions.sh](revisions.sh) to regenerate [revisions.docx](revisions.docx); open the result in Word and you will see all categories of marker in the review pane.

## Coverage matrix

| Section | Scope     | OOXML element                     | `revision.type` in `get`        | Created via                                                                          |
|--------:|-----------|-----------------------------------|---------------------------------|--------------------------------------------------------------------------------------|
|  1a     | run       | `w:ins`                           | `insertion`                     | `set /body/p[N]/r[M] --prop revision.type=ins --prop revision.author=…`              |
|  1b     | run       | `w:del`                           | `deletion`                      | `set /body/p[N]/r[M] --prop revision.type=del --prop revision.author=…`              |
|  1c     | run       | `w:rPrChange` (implicit)          | `formatChange`                  | `set /body/p[N]/r[M] --prop font.color=… --prop revision.author=…`                   |
|  1d     | run       | `w:rPrChange` (explicit)          | `formatChange`                  | `set /body/p[N]/r[M] --prop revision.type=format --prop italic=true --prop revision.author=…` |
|  2a     | paragraph | `w:ins` + `paragraphMarkInsertion`| `insertion` + `paragraphMarkInsertion` | `add /body --type paragraph --prop text=… --prop revision.author=…`           |
|  2b     | paragraph | `w:del` + `paragraphMarkDeletion` | `deletion`  + `paragraphMarkDeletion`  | `remove /body/p[N] --prop revision.author=…`                                  |
|  2c     | paragraph | `w:pPrChange`                     | `paragraphChange`               | `set /body/p[N] --prop align=… --prop revision.author=…`                             |
|  3      | run       | `w:moveFrom` / `w:moveTo` (paired)| `moveFrom` / `moveTo`           | two `set` calls with the **same** `revision.id=N`                                    |
|  4a     | table     | `w:tblPrChange`                   | `tableChange`                   | `set /body/tbl[N] --prop style=… --prop revision.author=…`                           |
|  4b     | row       | `w:trPrChange`                    | `rowChange`                     | `set /body/tbl[N]/tr[N] --prop height=… --prop revision.author=…`                    |
|  4c     | cell      | `w:tcPrChange` (+ cascades)       | `cellChange`                    | `set /body/tbl[N]/tr[N]/tc[N] --prop shd=… --prop revision.author=…`                 |
|  4d     | cell      | `w:tcPr/w:cellIns`                | `cellInsertion`                 | `add /body/tbl[N]/tr[N] --type cell --prop revision.author=…`                        |
|  4e     | cell      | `w:tcPr/w:cellDel`                | `cellDeletion`                  | `remove /body/tbl[N]/tr[N]/tc[N] --prop revision.author=…`                           |
|  4f     | row       | `w:trPr/w:ins`                    | `rowInsertion`                  | `add /body/tbl[N] --type row --prop revision.author=…`                               |
|  5      | section   | `w:sectPrChange`                  | `sectionChange`                 | `set /body/sectPr[N] --prop pageWidth=… --prop revision.author=…` (note: `/body/sectPr[N]`, **not** `/section[N]`) |
|  6a     | run       | default-author fallback           | `insertion` author=`OfficeCLI`  | `set … --prop revision.author=""` (empty string → `"OfficeCLI"`; only on `set`, not `add`) |
|  6b     | paragraph | explicit `revision.id`            | `insertion` id=9001             | `add … --prop revision.author=Grace --prop revision.id=9001`                         |

Side-effects worth knowing about (not separately addressable, but generated automatically):

- Cell `tcPrChange` cascades `w:tblPrExChange` per row + `w:tblGridChange` when grid columns mutate. Mac Word needs both for correct rendering; the handler emits them with unique revision ids — see commits `caf03ab2`, `b489826f`, `b271629f`.
- Move pairs always emit Range markers (`w:moveFromRangeStart/End`, `w:moveToRangeStart/End`) coalesced over contiguous runs — see `5e79cf8f`, `0f1f9a3a`.

## The two `revision.*` namespaces

Disjoint by design — mixing them throws.

- **Create** namespace: `revision.type` + `revision.author` + `revision.date` + `revision.id`, supplied on the host element (`add` or `set` on a run / paragraph / table / row / cell / section).
- **Action** namespace: `revision.action=accept|reject`, supplied on a `/revision[…]` selector or a native DOM path.
- A property change combined with `revision.author` (no `.type`, no `.action`) is the format-change path — writes the new value and captures the previous `rPr` / `pPr` / `tcPr` / `trPr` / `tblPr` / `sectPr` snapshot in the matching `*Change` element.
- Legacy `trackChange.*` keys were renamed to `revision.*` on this branch (`d9e812f3`). Bare `revision=…` is no longer accepted.

## Addressing markers

`query revision` returns one node per `w:ins` / `w:del` / `w:moveFrom` / `w:moveTo` / `w:*PrChange` / `w:*Ins` / `w:*Del`. The `Path` column is:

- `/revision[@id=N]` — **stable** across save (use this in scripts).
- `/revision[@id=N][@type=moveFrom|moveTo]` — single-end disambiguator for a move pair. `set /revision[@id=N]` without the `[@type=…]` segment still acts on **both** halves together (pair-wise accept/reject).
- `/revision[N]` — positional fallback for markers with no `w:id`. Indices shift after each accept/reject, so prefer `@id=`.

The `nativePath` column gives you the underlying DOM path (`/body/p[N]/ins[M]`, `/body/tbl[1]/tr[2]/tc[2]`, `/body/sectPr`, …), which works as a `set` target for `revision.action=…` too.

## Accept / reject syntax (all forms used in the demo)

```bash
# All
officecli set revisions.docx /revision --prop revision.action=accept
officecli set revisions.docx /revision --prop revision.action=reject

# By author / type
officecli set revisions.docx '/revision[@author=Alice]'  --prop revision.action=accept
officecli set revisions.docx '/revision[@type=ins]'      --prop revision.action=accept
officecli set revisions.docx '/revision[@type=del]'      --prop revision.action=reject

# By stable id (positional indices shift; ids do not)
officecli set revisions.docx '/revision[@id=42]'                 --prop revision.action=accept
officecli set revisions.docx '/revision[@id=42][@type=moveTo]'   --prop revision.action=reject

# Or via the native DOM path
officecli set revisions.docx '/body/p[2]/ins[1]'        --prop revision.action=accept
officecli set revisions.docx '/body/tbl[1]/tr[2]/tc[2]' --prop revision.action=accept
```

The script runs the accept/reject demo on a *temp copy* so the shipped `revisions.docx` keeps every marker in place for inspection.

## Notes

- `revision.date` accepts ISO-8601 (`2026-05-25T10:00:00Z`); omitted ⇒ `DateTime.UtcNow` at write time.
- `revision.author=""` on `set` falls back to `"OfficeCLI"`. On `add`, the empty author means "do not track" (the paragraph is added plain) — use `set` if you want the fallback behavior.
- `revision.id` is auto-allocated from the shared paraId pool when omitted. Supply explicitly to:
  - pair `moveFrom`/`moveTo` (must match),
  - get a deterministic id when downstream post-processing needs it (section 6b).
- The body's section properties path is **`/body/sectPr[N]`**, not `/section[N]`. `/section[N]` exists for mid-document sections (each `<w:p>` carrying a `pPr/sectPr`); the final body-level section lives on the body's trailing `sectPr` element.
- `revisionNumber` on `/` (document save counter) and `trackChanges` on `/settings` (the "Track Changes" mode toggle) are **separate** properties — neither creates per-edit markers.

See `officecli help docx revision` for the full property reference.
