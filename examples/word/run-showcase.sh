#!/bin/bash
# run-showcase.sh — exercise the docx run (character) property surface.
#
# Each paragraph demonstrates one family of run-level formatting. Most lines set
# the run formatting on the paragraph's implicit run via `add ... --type paragraph`;
# the super/subscript line uses explicit `--type run` children for mixed runs.
#
# Families: weight/style, underline variants + color, strike/dstrike, case
# (caps/smallCaps), vertical align (super/subscript), color/size/highlight,
# per-script fonts (latin/eastAsia/cs), text effects (emboss/imprint/outline/
# shadow/vanish), character spacing/kerning/position, and language tagging.
set -e

DOCX="$(dirname "$0")/run-showcase.docx"
echo "Building $DOCX ..."
rm -f "$DOCX"
officecli create "$DOCX"

heading() { officecli add "$DOCX" /body --type paragraph --prop "text=$1" --prop bold=true --prop size=14 --prop color=1F4E79 --prop spaceBefore=8pt; }

officecli add "$DOCX" /body --type paragraph --prop "text=Run / Character Formatting Showcase" --prop align=center --prop bold=true --prop size=20

# --- weight & style ---
heading "Weight & style"
officecli add "$DOCX" /body --type paragraph --prop "text=Bold text" --prop bold=true
officecli add "$DOCX" /body --type paragraph --prop "text=Italic text" --prop italic=true
officecli add "$DOCX" /body --type paragraph --prop "text=Bold + italic" --prop bold=true --prop italic=true

# --- underline variants + color ---
heading "Underline"
officecli add "$DOCX" /body --type paragraph --prop "text=single" --prop underline=single
officecli add "$DOCX" /body --type paragraph --prop "text=double" --prop underline=double
officecli add "$DOCX" /body --type paragraph --prop "text=thick" --prop underline=thick
officecli add "$DOCX" /body --type paragraph --prop "text=dotted" --prop underline=dotted
officecli add "$DOCX" /body --type paragraph --prop "text=wave (red)" --prop underline=wave --prop underline.color=FF0000

# --- strikethrough ---
heading "Strikethrough"
officecli add "$DOCX" /body --type paragraph --prop "text=single strike" --prop strike=true
officecli add "$DOCX" /body --type paragraph --prop "text=double strike" --prop dstrike=true

# --- case ---
heading "Case"
officecli add "$DOCX" /body --type paragraph --prop "text=all caps rendering" --prop caps=true
officecli add "$DOCX" /body --type paragraph --prop "text=small caps rendering" --prop smallcaps=true

# --- vertical align: super / subscript (mixed runs) ---
heading "Super / subscript"
officecli add "$DOCX" /body --type paragraph --prop "text=E = mc"
officecli add "$DOCX" "/body/p[last()]" --type run --prop "text=2" --prop superscript=true
officecli add "$DOCX" /body --type paragraph --prop "text=H"
officecli add "$DOCX" "/body/p[last()]" --type run --prop "text=2" --prop subscript=true
officecli add "$DOCX" "/body/p[last()]" --type run --prop "text=O"

# --- color / size / highlight ---
heading "Color, size, highlight"
officecli add "$DOCX" /body --type paragraph --prop "text=Red 16pt" --prop color=C00000 --prop size=16
officecli add "$DOCX" /body --type paragraph --prop "text=Highlighted" --prop highlight=yellow

# --- per-script fonts ---
heading "Per-script fonts"
officecli add "$DOCX" /body --type paragraph --prop "text=Latin Georgia + CJK 宋体" --prop font.latin=Georgia --prop font.eastAsia=SimSun --prop size=14

# --- text effects ---
heading "Text effects"
officecli add "$DOCX" /body --type paragraph --prop "text=emboss" --prop emboss=true
officecli add "$DOCX" /body --type paragraph --prop "text=imprint" --prop imprint=true
officecli add "$DOCX" /body --type paragraph --prop "text=outline" --prop outline=true
officecli add "$DOCX" /body --type paragraph --prop "text=shadow" --prop shadow=true

# --- character spacing / position ---
heading "Character spacing & position"
officecli add "$DOCX" /body --type paragraph --prop "text=expanded spacing" --prop charSpacing=2pt
officecli add "$DOCX" /body --type paragraph --prop "text=raised 3pt" --prop position=3pt

# --- language ---
heading "Language tag"
officecli add "$DOCX" /body --type paragraph --prop "text=Tagged en-US for spellcheck" --prop lang=en-US

officecli validate "$DOCX"
echo "Created: $DOCX"
