# Quick Reference Card

## 📋 Macro Usage

```
Alt+F11 → Run BuildPrintableDirectory()
```

## 📊 Output Sheets

| Sheet | Rows | Best For | Pages |
|-------|------|----------|-------|
| **A: PRINT-BY-NAME** | Sorted by name | Single/archive | ~200 |
| **B: PRINT-BY-UNIT** | Organized by units | Reference | ~150 |
| **B: PRINT-BY-UNIT-TOC** | Table of contents | Navigation | ~5 |
| **C: PRINT-BY-NAME-2COL** | Two-column compact | **Booklet** ⭐ | ~22 |
| **D: PRINT-BY-UNIT-2COL** | Two-column w/units | **Booklet** ⭐ | ~22 |

## 🎯 Recommended Workflow

1. **Use Sheet C or D** (two-column booklet formats)
2. **Print double-sided** (flip on short edge)
3. **Fold in half** (hamburger style)
4. **Center staple** (2-3 staples)
5. **Distribute** (~50-page booklets)

## 📐 Layout Specs

```
┌─ Left Column ─┬─ Gap ─┬─ Right Column ─┐
│ 16" wide      │ 1.5"  │ 16" wide       │
│ 38 rows       │       │ 38 rows        │
│ 12pt Calibri  │       │ 12pt Calibri   │
│ 16.5pt height │       │ 16.5pt height  │
└───────────────┴───────┴────────────────┘
```

## 🔧 Key Files

| File | Edit For |
|------|----------|
| `modular_core.bas` | Sheet names, page prefixes |
| `modular_two_column.bas` | Font size, row count, column width |
| `modular_helpers.bas` | Text processing logic |

## ⚙️ Configuration

Change in `modular_core.bas`:

```vba
' Font size for seniors
Private Const FONT_SIZE_BODY As Double = 12  ' ← Change here (11-14 typical)

' Rows per column
Private Const TWO_COL_ROWS_PER_COLUMN As Long = 38  ' ← Change here
```

## 🖨️ Print Settings

- **Orientation**: Portrait
- **Paper**: Letter (8.5×11")
- **Two-sided**: ✓ Yes
- **Scale**: 100% (optimized)
- **Margins**: 0.5" all (pre-set)

## 📱 Senior-Friendly Features

✓ Large font (12pt)
✓ Generous spacing (16.5pt rows)
✓ High contrast (black/white)
✓ Clear headers (repeating)
✓ Simple layout (no clutter)

## 📁 File Organization

```
directory_formatter/
├── vba/
│   ├── modular_core.bas
│   ├── modular_parsing.bas
│   ├── modular_sorting.bas
│   ├── modular_layout.bas
│   ├── modular_two_column.bas
│   └── modular_helpers.bas
├── README.md
├── COMPLETE_SOLUTION_SUMMARY.md
└── [.csv & .xlsx ignored]
```

## 🚀 Quick Start

1. **Import** → Alt+F11 → Import 6 .bas files
2. **Paste** → CSV into "PASTE-HERE" sheet
3. **Run** → BuildPrintableDirectory()
4. **Print** → Sheet C or D, double-sided
5. **Fold** → Create booklets
6. **Done!** → ~50-page professional booklets

## 🆘 Troubleshooting

| Problem | Solution |
|---------|----------|
| Blank output | Check PASTE-HERE sheet exists |
| No two-column sheets | Verify single-column sheets have data |
| Font too small | Edit `FONT_SIZE_BODY` in modular_two_column.bas |
| Page breaks wrong | Check `START_EACH_UNIT_ON_NEW_PAGE` setting |
| Headers missing | Sheets have freeze panes set (correct) |

## 📞 Support Resources

- **README.md** — Full instructions
- **COMPLETE_SOLUTION_SUMMARY.md** — Detailed overview
- **TWO_COLUMN_IMPLEMENTATION.md** — Technical details
- **TWO_COLUMN_LAYOUT_GUIDE.md** — Visual diagrams

---

**Status**: ✅ Ready for production
**Sheets**: 5 total (2 single-col + 2 booklet + 1 TOC)
**Booklet Size**: ~50 pages (compact & senior-friendly)
**Time to Print**: ~15 minutes
