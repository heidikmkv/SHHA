# Two-Column Layout Visual Guide

## Page Layout (8.5x11" Portrait)

```
┌────────────────────────────────────────────────────────────┐
│                      0.5" margin                           │
│ ┌──────────────────────────────────────────────────────┐  │
│ │                                                      │  │
│ │  LEFT COLUMN             │      │ RIGHT COLUMN      │  │
│ │  ─────────────────────── │ GAP  │ ──────────────────│  │
│ │  Last Name │ First │Phone│      │ Last │ First │Phone│  │
│ │  Smith     │ John  │555- │      │ Jones│ Mary  │555- │  │
│ │  Brown     │ Jane  │1234 │      │ White│ Tom   │5678 │  │
│ │  ...       │ ...   │ ... │      │ ...  │ ...   │ ... │  │
│ │  [38 rows] │       │     │      │[38 rows]       │     │  │
│ │            │       │     │      │                │     │  │
│ │  ──────────────────────── ────── ──────────────────│  │
│ │  [Next page section starts] [Auto page break]      │  │
│ │                                                      │  │
│ └──────────────────────────────────────────────────────┘  │
│                                                            │
│                    Footer: C-1, C-2, etc.                │
└────────────────────────────────────────────────────────────┘
```

## Column Widths

| Column | Width | Content |
|--------|-------|---------|
| A | 16 | Last Name |
| B | 16 | First Name |
| C | 13 | Phone # |
| D | 7  | Street # |
| E | 20 | Street Name |
| F | 5  | Mbr? (Yes/No) |
| G | 1.5 | **SPACER** |
| H | 1.5 | **SPACER** |
| I | 16 | Last Name (Right) |
| J | 16 | First Name (Right) |
| K | 13 | Phone # (Right) |
| L | 7  | Street # (Right) |
| M | 20 | Street Name (Right) |
| N | 5  | Mbr? (Right) |

## Row Structure

```
Row 1:     Headers (bold, 13pt, gray background)
           ├─ Last Name  │ First Name  │ Phone # │ ... (Left)
           │             │             │         │
           └─ Last Name  │ First Name  │ Phone # │ ... (Right)

Rows 2-39: Data rows (12pt, 16.5pt row height)
           [38 entries total = left column complete]

Row 40:    [Repeat for next section or page break]
```

## Typical Page Flow

### Sheet: PRINT-BY-NAME-2COL
```
Physical Page 1 (double-sided)
├─ Front (Top view = Rows 1-40 left, rows 1-40 right)
│  ├─ Headers (Smith, Jones, etc.)
│  └─ ~76 names visible
│
└─ Back (bottom view = Rows 41-80 left, rows 41-80 right)
   ├─ Headers repeat (Brown, White, etc.)
   └─ ~76 more names visible

Result: 152 names per physical page (double-sided)
For 3300 names: 3300 ÷ 152 = ~22 pages
Folded booklet: ~44 pages
```

### Sheet: PRINT-BY-UNIT-2COL
```
Same layout but with unit section headers:

Row 1:     [Headers]
Row 2:     Unit: South District 10  (merged A:F, bold, gray bg)
Row 3-15:  [Data rows for South District 10]
Row 16:    Unit: North District 5
Row 17-30: [Data rows for North District 5]
...
(Page break when section won't fit in current column)
```

## Font & Spacing Details

| Element | Size | Font | Color | Background |
|---------|------|------|-------|------------|
| Header row | 13pt | Calibri | Black | RGB(200, 200, 200) |
| Data rows | 12pt | Calibri | Black | White |
| Row height | 16.5pt | — | — | — |
| Unit headers | 13pt | Calibri Bold | Black | RGB(220, 220, 220) |

## Alignment

| Column | Alignment | Reasoning |
|--------|-----------|-----------|
| Last Name | Left | Names easier to scan left-aligned |
| First Name | Left | Standard name format |
| Phone | Center | Easier to spot phone numbers |
| Street # | Center | Numbers easier to locate |
| Street | Left | Address block better left-aligned |
| Member? | Center | Yes/No answers quick to find |

## Page Numbering

- **A series**: PRINT-BY-NAME (single-column) → A-1, A-2, A-3...
- **B series**: PRINT-BY-UNIT (single-column) → B-1, B-2, B-3...
- **C series**: PRINT-BY-NAME-2COL (booklet) → C-1, C-2, C-3...
- **D series**: PRINT-BY-UNIT-2COL (booklet) → D-1, D-2, D-3...

## Senior Readability Features

✓ **Large font** (12pt vs 11pt in single-column)
✓ **Generous spacing** (16.5pt row height)
✓ **High contrast** (black text on white)
✓ **Clear headers** (repeating on each page section)
✓ **Simple layout** (no merging or complex formatting)
✓ **Calibri font** (clean, easy to read)
✓ **Appropriate color** (subtle gray headers, not bright)

## Printing Recommendations

1. **Color**: Black & white sufficient (gray headers print fine)
2. **Density**: 600 DPI minimum recommended
3. **Paper**: 20lb standard office paper
4. **Binding**: Center staple (2-3 staples) for ~50-page booklets
5. **Cover**: Optional cardstock cover (11x8.5") for durability
6. **Margins**: Already set (0.5" all sides) - no adjustment needed
