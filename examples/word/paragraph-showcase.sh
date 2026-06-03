#!/bin/bash
# paragraph-showcase.sh — exercise the docx paragraph property surface.
#
# Sections: alignment, indentation, spacing, pagination flags, paragraph-level
# run formatting (applied to every run), shading, the paragraph-mark run
# properties (markRPr.* — formatting of the ¶ glyph itself), and outline level.
#
# Note: paragraph-level `bold` etc. apply to all runs in the paragraph and read
# back on the paragraph; `markRPr.bold` formats only the paragraph mark and is
# distinct. Both are settable + gettable.
set -e

DOCX="$(dirname "$0")/paragraph-showcase.docx"
echo "Building $DOCX ..."
rm -f "$DOCX"
officecli create "$DOCX"

heading() { officecli add "$DOCX" /body --type paragraph --prop "text=$1" --prop bold=true --prop size=14 --prop color=1F4E79 --prop spaceBefore=10pt; }

officecli add "$DOCX" /body --type paragraph --prop "text=Paragraph Formatting Showcase" --prop align=center --prop bold=true --prop size=20

# --- alignment ---
heading "Alignment"
officecli add "$DOCX" /body --type paragraph --prop "text=Left aligned (default)" --prop align=left
officecli add "$DOCX" /body --type paragraph --prop "text=Center aligned" --prop align=center
officecli add "$DOCX" /body --type paragraph --prop "text=Right aligned" --prop align=right
officecli add "$DOCX" /body --type paragraph --prop "text=Justified text stretched edge to edge across the full measure of the line so both margins align." --prop align=both

# --- indentation ---
heading "Indentation"
officecli add "$DOCX" /body --type paragraph --prop "text=Left indent 1cm" --prop indent=1cm
officecli add "$DOCX" /body --type paragraph --prop "text=Right indent 2cm so the right edge pulls in." --prop rightIndent=2cm
officecli add "$DOCX" /body --type paragraph --prop "text=First-line indent — only the first line is pushed in from the left margin." --prop firstLineIndent=1cm
officecli add "$DOCX" /body --type paragraph --prop "text=Hanging indent — the first line hangs left while the rest of the paragraph is indented." --prop indent=1cm --prop hangingIndent=1cm

# --- spacing ---
heading "Spacing"
officecli add "$DOCX" /body --type paragraph --prop "text=Space before 18pt, after 6pt" --prop spaceBefore=18pt --prop spaceAfter=6pt
officecli add "$DOCX" /body --type paragraph --prop "text=Line spacing 1.5x across a longer paragraph that wraps so the extra leading between wrapped lines is visible." --prop lineSpacing=1.5x

# --- pagination flags ---
heading "Pagination flags"
officecli add "$DOCX" /body --type paragraph --prop "text=keepNext — stays with the following paragraph" --prop keepNext=true
officecli add "$DOCX" /body --type paragraph --prop "text=keepLines — lines stay together, never split across pages" --prop keepLines=true
officecli add "$DOCX" /body --type paragraph --prop "text=widowControl on" --prop widowControl=true

# --- paragraph-level run formatting (applies to all runs) ---
heading "Paragraph-level run formatting"
officecli add "$DOCX" /body --type paragraph --prop "text=Whole paragraph bold + red + 13pt" --prop bold=true --prop color=C00000 --prop size=13
officecli add "$DOCX" /body --type paragraph --prop "text=Whole paragraph italic + highlighted" --prop italic=true --prop highlight=yellow

# --- shading ---
heading "Shading"
officecli add "$DOCX" /body --type paragraph --prop "text=Light gray paragraph shading" --prop shading.fill=D9D9D9
officecli add "$DOCX" /body --type paragraph --prop "text=Pale blue shading" --prop shading.fill=DDEBF7

# --- paragraph-mark run props (the pilcrow itself) ---
heading "Paragraph-mark formatting (markRPr)"
officecli add "$DOCX" /body --type paragraph --prop "text=The mark glyph is bold+red (distinct from run text)" --prop markRPr.bold=true --prop markRPr.color=C00000

# --- outline level ---
heading "Outline level"
officecli add "$DOCX" /body --type paragraph --prop "text=Outline level 1 (shows in document map)" --prop outlineLvl=1

officecli validate "$DOCX"
echo "Created: $DOCX"
