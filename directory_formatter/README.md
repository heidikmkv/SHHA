# Directory Formatter - Modular VBA Refactor

This folder contains a refactored Excel directory printer macro split into modular .bas files for better maintainability and version control.

## File Structure

- **modular_core.bas** - Main entry point (`BuildPrintableDirectory()`) and configuration
- **modular_parsing.bas** - Name/phone expansion with Resident flag handling
- **modular_sorting.bas** - Sort routines for BY-NAME and BY-UNIT layouts
- **modular_layout.bas** - Single-column sheet building and formatting functions
- **modular_two_column.bas** - Two-column senior-friendly booklet formatting
- **.gitignore** - Excludes CSV and XLSX files (sensitive homeowner data)

## How to Import into Excel

### Option 1: Manual Import via VB Editor (Recommended)

1. Open your Excel workbook (`.xlsm`)
2. Press **Alt + F11** to open Visual Basic Editor
3. In the Project Explorer on the left, right-click your project name
4. Select **File** → **Import File**
5. Select the first file: `modular_core.bas`
6. Repeat steps 4-5 for each remaining .bas file:
   - `modular_parsing.bas`
   - `modular_sorting.bas`
   - `modular_layout.bas`
   - `modular_helpers.bas`
   - `modular_two_column.bas`
7. Click **Tools** → **References** and ensure Excel Object Library is checked
8. Close the VB Editor (Alt + Q)

### Option 2: Macro to Auto-Import (Advanced)

You can create a helper macro in your workbook to import these files programmatically, but manual import is more straightforward for most users.

## Usage

1. Paste your website CSV export into the **PASTE-HERE** sheet (starting at cell A1)
2. Run `BuildPrintableDirectory()` macro
3. The macro will generate **five output sheets**:

### Single-Column Formats (Original Style)
   - **PRINT-BY-NAME** (A-#) - Sorted by last name, omits "Resident" entries, ~200 pages
   - **PRINT-BY-UNIT** (B-#) - Organized by HOA Unit/District, includes "Resident" entries
   - **PRINT-BY-UNIT-TOC** (B-#) - Table of contents for PRINT-BY-UNIT

### Two-Column Booklet Formats (Senior-Friendly, ~50 pages)
   - **PRINT-BY-NAME-2COL** (C-#) - Two-column by-name layout, ideal for booklet printing
   - **PRINT-BY-UNIT-2COL** (D-#) - Two-column by-unit layout, ideal for booklet printing

#### Two-Column Features:
- **Senior-optimized**: 12pt Calibri body text (larger than single-column sheets)
- **Generous spacing**: 16.5pt row height for easy reading
- **Clear separation**: Visual gap between left and right columns
- **Compact printing**: ~4 pages per 8.5x11" sheet (double-sided) instead of 1 page
- **Booklet ready**: Print double-sided, fold in half to create booklets (~50 pages vs ~200)
- **Intelligent page breaks**: Headers repeat, unit sections honored
- **Footer pages**: Labeled C-1, C-2, D-1, D-2 for tracking

## Configuration

Edit settings at the top of `modular_core.bas`:

```vba
Private Const INPUT_SHEET As String = "PASTE-HERE"
Private Const OUT_BY_NAME As String = "PRINT-BY-NAME"
Private Const OUT_BY_UNIT As String = "PRINT-BY-UNIT"
Private Const OUT_BY_UNIT_TOC As String = "PRINT-BY-UNIT-TOC"
Private Const OUT_BY_NAME_2COL As String = "PRINT-BY-NAME-2COL"
Private Const OUT_BY_UNIT_2COL As String = "PRINT-BY-UNIT-2COL"

Private Const PAGE_PREFIX_BY_NAME As String = "A"
Private Const PAGE_PREFIX_BY_UNIT As String = "B"
Private Const PAGE_PREFIX_TOC As String = "B"
Private Const PAGE_PREFIX_BY_NAME_2COL As String = "C"
Private Const PAGE_PREFIX_BY_UNIT_2COL As String = "D"

Private Const START_EACH_UNIT_ON_NEW_PAGE As Boolean = True
```

## Printing & Booklet Creation

### Recommended Workflow:
1. Use **PRINT-BY-NAME-2COL** or **PRINT-BY-UNIT-2COL** for final printing
2. Print double-sided (flip on short edge for binding)
3. Fold in half to create a booklet
4. The page numbering (C-1, C-2, etc.) helps organize pages

### Example for ~50-page Booklet:
- Sheet: PRINT-BY-NAME-2COL (38 data rows × 2 columns = ~76 entries per page view)
- Page count: ~22 physical pages (4 views per page × 2 sides) = ~44 pages when folded
- Print setting: "Fit to 1 page wide" (already configured)

## Data Requirements

Your CSV must include these columns (in any order):
- `Directory Names` - Full names (supports "First & First Last" and newline-separated)
- `Directory Phone Numbers` - Phone numbers (newline-separated)
- `Number` - Street number
- `Street` - Street name
- `Unit` - Apartment/unit letter/number (optional)
- `HOA Unit` - Grouping unit (optional; used for sorting)
- `District` - Alternative grouping unit (fallback if HOA Unit missing)
- `Is Member` - YES/NO or 1/0 (required)

## Name Parsing Rules

- **Standard**: "John Smith" → First: John, Last: Smith
- **Ampersand**: "John & Jane Smith" → Two rows (John Smith, Jane Smith)
- **Newline-separated**: Multiple names on same address
- **Resident**: Special marker for addresses with unnamed residents

## Editing & Development

To modify the code:

1. Edit the `.bas` files in your preferred text editor (VS Code recommended)
2. Test changes by re-importing the modified `.bas` file into Excel
3. To replace a module: Delete it in VB Editor, then import the updated file
4. Commit changes to git (CSV/XLSX files are ignored by .gitignore)

## Troubleshooting

### "Missing required headers" error
- Verify your CSV has the exact column names listed above
- Check that data is pasted starting at cell A1

### "No entries produced" message
- Check the input sheet name matches configuration
- Verify addresses aren't blank (Street + Number)

### Macro doesn't appear in macro list
- Ensure module names don't start with underscore
- All files should be imported as standard modules (not class modules)

## Git Workflow

```bash
git add modular_*.bas .gitignore README.md
git commit -m "Refactor: split VBA code into modular files"
# Never commit:
git add --ignore-errors *.csv *.xlsx
```

## License

Internal use only - contains sensitive homeowner directory data.
