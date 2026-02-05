# Two-Column Booklet Feature - Implementation Summary

## What Was Added

### New Module: `modular_two_column.bas`
A complete two-column layout formatter optimized for senior-friendly reading:

**Key Functions:**
- `BuildTwoColumnByName()` - Converts PRINT-BY-NAME to compact two-column format
- `BuildTwoColumnByUnit()` - Converts PRINT-BY-UNIT to compact two-column format
- Intelligent page breaking and header repetition
- Senior-optimized typography and spacing

**Layout Details:**
```
┌─────────────────────┬─────┬──────────────────────┐
│ LEFT COLUMN         │ GAP │ RIGHT COLUMN         │
│ (cols A-F)          │ (H) │ (cols I-N)           │
│ ▪ Last Name         │     │ ▪ Last Name          │
│ ▪ First Name        │     │ ▪ First Name         │
│ ▪ Phone #           │     │ ▪ Phone #            │
│ ▪ Street #          │     │ ▪ Street #           │
│ ▪ Street            │     │ ▪ Street             │
│ ▪ Member?           │     │ ▪ Member?            │
│                     │     │                      │
│ (38 rows per col)   │     │ (38 rows per col)    │
└─────────────────────┴─────┴──────────────────────┘
```

### Modified Module: `modular_core.bas`
- Added constants for new output sheets:
  - `OUT_BY_NAME_2COL` → "PRINT-BY-NAME-2COL" (Sheet C)
  - `OUT_BY_UNIT_2COL` → "PRINT-BY-UNIT-2COL" (Sheet D)
- Added page prefixes: C-# and D-#
- Generates both sheets automatically after single-column variants
- Calls `BuildTwoColumnByName()` and `BuildTwoColumnByUnit()`

### Updated README.md
- Documents new two-column sheets
- Explains booklet printing workflow
- Provides page count estimates
- Lists senior-friendly features

## Features

### Typography (Senior-Optimized)
- **Font**: Calibri 12pt (body), 13pt (headers)
- **Row Height**: 16.5pt for generous spacing
- **Header Row Height**: 20pt
- **Bold Headers**: Dark gray (RGB 200,200,200) background

### Layout Benefits
- **50% Page Reduction**: ~200 pages → ~50 pages double-sided booklet
- **Column Separation**: Visual gap between left/right columns
- **Clear Headers**: Repeat on every page section
- **Intelligent Breaks**: Respects unit groupings, avoids orphaned entries
- **Professional Footer**: Page numbers (C-1, C-2, etc.) for reference

### Print Optimization
- **Margins**: 0.5" sides, 0.5" top/bottom
- **Paper Size**: 8.5x11" (Letter)
- **Orientation**: Portrait
- **Fit to Page**: Optimized for single-page width
- **Title Rows**: Headers freeze and print on all pages

## Usage

After importing all modules, run the macro normally:
1. Paste CSV into "PASTE-HERE" sheet
2. Run `BuildPrintableDirectory()` macro
3. Five sheets are generated:
   - **A: PRINT-BY-NAME** (original single-column)
   - **B: PRINT-BY-UNIT** (original single-column)
   - **B: PRINT-BY-UNIT-TOC** (table of contents)
   - **C: PRINT-BY-NAME-2COL** ← Use for booklet
   - **D: PRINT-BY-UNIT-2COL** ← Use for booklet

## Printing Instructions

### For Booklet Creation:
1. Open "PRINT-BY-NAME-2COL" or "PRINT-BY-UNIT-2COL"
2. **Print Settings**:
   - Layout: Portrait
   - Scale: 100% (already optimized)
   - Two-sided: Yes (flip on short edge)
   - Paper type: Standard letter
3. **After Printing**:
   - Stack pages in order
   - Fold in half
   - Staple in center
4. Result: Professional ~50-page booklet

### For Standard (Non-Booklet) Printing:
- Use original "PRINT-BY-NAME" or "PRINT-BY-UNIT" sheets
- Print single-sided or double-sided as preferred
- No folding required

## Technical Details

### Data Flow
```
CSV Input
  ↓
modular_core.bs (parsing, buffering)
  ↓
modular_sorting.bas (sorts single-column data)
  ├→ PRINT-BY-NAME (sheet A)
  ├→ PRINT-BY-UNIT (sheet B)
  └→ PRINT-BY-UNIT-TOC (sheet B)
  ↓
modular_two_column.bas (converts to 2-col format)
  ├→ PRINT-BY-NAME-2COL (sheet C)
  └→ PRINT-BY-UNIT-2COL (sheet D)
```

### Row Calculation
- **Per Column**: 38 data rows
- **Per Sheet View**: 38 × 2 = 76 entries visible
- **Per Physical Page** (double-sided): 4 sheet views = 152 entries per page
- **For 3300 names**: 3300 ÷ 152 = ~22 pages × 2 sides = ~44-page booklet

## Testing Notes

To verify the output quality:
1. Check that fonts are crisp and readable (no overlapping)
2. Verify column headers print on every page
3. Confirm page numbers (C-1, C-2, etc.) appear in footer
4. Print a sample double-sided to verify fold alignment

## Future Enhancements (Optional)

- Add zebra striping (alternating row colors) for additional readability
- Configurable rows-per-column (currently fixed at 38)
- Option to include blank spacing between unit sections
- Color-coded unit headers for quick navigation
