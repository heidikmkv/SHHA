#!/bin/bash

#=========================================================
# VBA Monolith Builder Script
#=========================================================
# Concatenates all modular .bas files into a single file
# Usage: ./build_monolith.sh
#=========================================================

OUTPUT_FILE="directory_formatter/directory_formatter_combined.bas"
TEMP_FILE="/tmp/vba_monolith_$$.bas"

echo "Building monolith VBA file..."

# Start with Option Explicit
echo "Option Explicit" > "$TEMP_FILE"
echo "" >> "$TEMP_FILE"

# Array of files in order of import
FILES=(
    "directory_formatter/vba/modular_core.bas"
    "directory_formatter/vba/modular_parsing.bas"
    "directory_formatter/vba/modular_sorting.bas"
    "directory_formatter/vba/modular_layout.bas"
    "directory_formatter/vba/modular_two_column.bas"
    "directory_formatter/vba/modular_helpers.bas"
)

# Process each file
for file in "${FILES[@]}"; do
    if [ ! -f "$file" ]; then
        echo "❌ Error: File not found: $file"
        exit 1
    fi
    
    # Extract module name from filename
    module_name=$(basename "$file" .bas)
    
    echo "✓ Processing: $module_name"
    
    # Add section header
    echo "" >> "$TEMP_FILE"
    echo "'==========================================================" >> "$TEMP_FILE"
    echo "' MODULE: $module_name" >> "$TEMP_FILE"
    echo "'==========================================================" >> "$TEMP_FILE"
    echo "" >> "$TEMP_FILE"
    
    # Add file contents, skipping duplicate "Option Explicit" lines
    skip_option_explicit=true
    while IFS= read -r line; do
        # Skip the first "Option Explicit" of each file (we already have one at top)
        if [[ "$line" == "Option Explicit" ]]; then
            if [ "$skip_option_explicit" = true ]; then
                skip_option_explicit=false
                continue
            fi
        fi
        
        # Skip empty lines at start of file
        if [[ -z "$line" ]] && [ "$skip_option_explicit" = true ]; then
            continue
        fi
        
        # Change "Private Function" to "Public Function" (except in helpers)
        if [[ "$line" =~ ^Private\ (Function|Sub) ]]; then
            # Only make public the main functions that need to be called
            if [[ "$module_name" == "modular_core" || "$module_name" == "modular_two_column" ]]; then
                line="${line//Private/Public}"
            fi
        fi
        
        echo "$line" >> "$TEMP_FILE"
    done < "$file"
done

# Move temp file to output location
mv "$TEMP_FILE" "$OUTPUT_FILE"

echo ""
echo "✅ SUCCESS: Monolith created at: $OUTPUT_FILE"
echo ""
echo "📋 Instructions:"
echo "1. Open the new file: $OUTPUT_FILE"
echo "2. Select ALL (Ctrl+A)"
echo "3. Copy (Ctrl+C)"
echo "4. In Excel VB Editor (Alt+F11):"
echo "   - Right-click project → Insert Module"
echo "   - Paste all code"
echo "   - Run BuildPrintableDirectory()"
echo ""
echo "📊 File stats:"
wc -l "$OUTPUT_FILE"
echo ""
