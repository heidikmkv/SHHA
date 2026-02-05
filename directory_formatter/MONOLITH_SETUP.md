# Quick Monolith Setup Guide

## What You Have

A single combined VBA file: `directory_formatter_combined.bas` (1,268 lines)

This contains all 6 modules merged together for easy copy-paste into Excel.

## How to Use

### Option A: Copy-Paste (Fastest)

1. **Open the file**
   ```
   directory_formatter/directory_formatter_combined.bas
   ```

2. **Select ALL (Ctrl+A) and Copy (Ctrl+C)**

3. **In Excel**:
   - Press **Alt+F11** (Open VB Editor)
   - Right-click project → **Insert Module**
   - Paste the code

4. **Run the macro**
   - In VB Editor: **Alt+F5** (or Tools → Macros → BuildPrintableDirectory)
   - Or press the macro button in Excel

### Option B: Import Individual Modules (Original Way)

If you prefer the modular approach (cleaner):

1. Go to `/vba/` folder
2. Import files in order:
   - `modular_core.bas`
   - `modular_parsing.bas`
   - `modular_sorting.bas`
   - `modular_layout.bas`
   - `modular_two_column.bas`
   - `modular_helpers.bas`

Both approaches give identical results.

## Regenerate the Monolith

If you edit the individual modules and want to rebuild the combined file:

```bash
cd directory_formatter
./build_monolith.sh
```

This recreates `directory_formatter_combined.bas` with all changes.

## File Structure

```
directory_formatter/
├── build_monolith.sh                      ← Script to rebuild monolith
├── directory_formatter_combined.bas       ← Single file for copy-paste ⭐
├── vba/
│   ├── modular_core.bas
│   ├── modular_parsing.bas
│   ├── modular_sorting.bas
│   ├── modular_layout.bas
│   ├── modular_two_column.bas
│   └── modular_helpers.bas
└── [docs + config files]
```

## Testing Checklist

- [ ] Paste code into Excel VB Editor
- [ ] No syntax errors (red squiggly lines)
- [ ] Create "PASTE-HERE" sheet
- [ ] Paste CSV starting at A1
- [ ] Run BuildPrintableDirectory()
- [ ] 5 sheets appear (A, B, B-TOC, C, D)
- [ ] Sheet C/D have two-column layout
- [ ] Headers repeat on page 2
- [ ] Page numbers in footer

## Troubleshooting

**"Sub not found"**
→ Make sure you're in a regular module (not class)

**Code highlighted in red**
→ Check for missing apostrophes or unmatched quotes

**Blank output**
→ Verify "PASTE-HERE" sheet exists and CSV is there

## Tips

- The monolith is read-only for git (just for testing)
- Keep your edits in the `/vba/` modules
- Run `build_monolith.sh` to update the combined file
- Both approaches work identically in Excel

---

**Size**: 1,268 lines (vs 988 in original, +280 for two-column feature)
**Status**: ✅ Ready for paste-and-test
**Recommended**: Use monolith for quick testing, modules for long-term maintenance
