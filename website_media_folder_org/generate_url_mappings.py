#!/usr/bin/env python3
"""
URL Mapping Generator for Moved/Renamed Files
Compares before/after FTP inventories to generate URL redirect mappings.
"""

import json
import csv
from pathlib import Path
from typing import Dict, List, Tuple
from datetime import datetime
import os


class URLMappingGenerator:
    """Generate URL mappings for moved/renamed files based on MD5 hash matching."""
    
    def __init__(self, base_url: str = ""):
        """
        Initialize URL mapping generator.
        
        Args:
            base_url: Base URL for the website (e.g., 'https://example.com')
        """
        self.base_url = base_url.rstrip('/')
        self.before_inventory: Dict[str, Dict] = {}
        self.after_inventory: Dict[str, Dict] = {}
        self.mappings: List[Dict] = []
    
    def load_inventory_json(self, filepath: str) -> Dict[str, Dict]:
        """
        Load inventory from JSON file and index by MD5 hash.
        
        Args:
            filepath: Path to JSON inventory file
            
        Returns:
            Dictionary mapping MD5 hash to file info
        """
        with open(filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
        
        inventory = {}
        for file_info in data.get('files', []):
            md5 = file_info.get('md5')
            if md5 and md5 != 'ERROR':
                inventory[md5] = file_info
        
        return inventory
    
    def load_inventory_csv(self, filepath: str) -> Dict[str, Dict]:
        """
        Load inventory from CSV file and index by MD5 hash.
        
        Args:
            filepath: Path to CSV inventory file
            
        Returns:
            Dictionary mapping MD5 hash to file info
        """
        inventory = {}
        
        with open(filepath, 'r', encoding='utf-8') as f:
            reader = csv.DictReader(f)
            for row in reader:
                md5 = row.get('md5')
                if md5 and md5 != 'ERROR':
                    inventory[md5] = row
        
        return inventory
    
    def generate_mappings(self, before_file: str, after_file: str):
        """
        Generate URL mappings by comparing before and after inventories.
        
        Args:
            before_file: Path to inventory taken before reorganization
            after_file: Path to inventory taken after reorganization
        """
        print(f"\n{'='*70}")
        print("Generating URL Mappings")
        print(f"{'='*70}\n")
        
        # Determine file format and load inventories
        if before_file.endswith('.json'):
            self.before_inventory = self.load_inventory_json(before_file)
        else:
            self.before_inventory = self.load_inventory_csv(before_file)
        
        if after_file.endswith('.json'):
            self.after_inventory = self.load_inventory_json(after_file)
        else:
            self.after_inventory = self.load_inventory_csv(after_file)
        
        print(f"✓ Loaded BEFORE inventory: {len(self.before_inventory)} files")
        print(f"✓ Loaded AFTER inventory: {len(self.after_inventory)} files\n")
        
        # Find moved/renamed files
        moved_count = 0
        unchanged_count = 0
        new_count = 0
        deleted_count = 0
        
        # Track which MD5s we've seen in the after inventory
        after_md5s = set(self.after_inventory.keys())
        before_md5s = set(self.before_inventory.keys())
        
        # Find moved files (same MD5, different path)
        for md5 in before_md5s:
            before_path = self.before_inventory[md5]['path']
            
            if md5 in after_md5s:
                after_path = self.after_inventory[md5]['path']
                
                if before_path != after_path:
                    # File was moved/renamed
                    self.mappings.append({
                        'old_path': before_path,
                        'new_path': after_path,
                        'old_url': self._path_to_url(before_path),
                        'new_url': self._path_to_url(after_path),
                        'md5': md5,
                        'filename': self.after_inventory[md5]['name'],
                        'size': self.after_inventory[md5]['size'],
                        'status': 'moved'
                    })
                    moved_count += 1
                else:
                    # File unchanged
                    unchanged_count += 1
            else:
                # File was deleted
                deleted_count += 1
        
        # Count new files (in after but not in before)
        new_count = len(after_md5s - before_md5s)
        
        print(f"Analysis Results:")
        print(f"  Moved/Renamed: {moved_count}")
        print(f"  Unchanged:     {unchanged_count}")
        print(f"  New files:     {new_count}")
        print(f"  Deleted:       {deleted_count}")
        print(f"\n{'='*70}\n")
        
        return self.mappings
    
    def _path_to_url(self, path: str) -> str:
        """
        Convert FTP path to URL.
        
        Args:
            path: FTP file path
            
        Returns:
            Full URL
        """
        # Remove leading slash and create URL
        clean_path = path.lstrip('/')
        if self.base_url:
            return f"{self.base_url}/{clean_path}"
        return f"/{clean_path}"
    
    def save_csv(self, output_file: str = "url_mappings.csv"):
        """Save URL mappings to CSV file."""
        if not self.mappings:
            print("No mappings to save (no files were moved)")
            return
        
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            fieldnames = ['old_url', 'new_url', 'old_path', 'new_path', 'filename', 'size', 'md5', 'status']
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(self.mappings)
        
        print(f"✓ CSV mappings saved to: {output_file}")
    
    def save_json(self, output_file: str = "url_mappings.json"):
        """Save URL mappings to JSON file."""
        if not self.mappings:
            print("No mappings to save (no files were moved)")
            return
        
        output = {
            'generated_at': datetime.now().isoformat(),
            'total_mappings': len(self.mappings),
            'base_url': self.base_url,
            'mappings': self.mappings
        }
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(output, f, indent=2)
        
        print(f"✓ JSON mappings saved to: {output_file}")
    
    def save_htaccess(self, output_file: str = "redirects.htaccess"):
        """
        Generate Apache .htaccess redirect rules.
        
        Args:
            output_file: Output filename for redirect rules
        """
        if not self.mappings:
            print("No mappings to save (no files were moved)")
            return
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("# Generated URL Redirects\n")
            f.write(f"# Generated: {datetime.now().isoformat()}\n")
            f.write(f"# Total redirects: {len(self.mappings)}\n\n")
            f.write("RewriteEngine On\n\n")
            
            for mapping in self.mappings:
                old_path = mapping['old_path'].lstrip('/')
                new_path = mapping['new_path'].lstrip('/')
                f.write(f"# {mapping['filename']}\n")
                f.write(f"RewriteRule ^{old_path}$ /{new_path} [R=301,L]\n\n")
        
        print(f"✓ Apache redirects saved to: {output_file}")
    
    def save_nginx(self, output_file: str = "redirects.nginx.conf"):
        """
        Generate Nginx redirect rules.
        
        Args:
            output_file: Output filename for redirect rules
        """
        if not self.mappings:
            print("No mappings to save (no files were moved)")
            return
        
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write("# Generated URL Redirects\n")
            f.write(f"# Generated: {datetime.now().isoformat()}\n")
            f.write(f"# Total redirects: {len(self.mappings)}\n\n")
            
            for mapping in self.mappings:
                old_path = mapping['old_path']
                new_path = mapping['new_path']
                f.write(f"# {mapping['filename']}\n")
                f.write(f"location = {old_path} {{\n")
                f.write(f"    return 301 {new_path};\n")
                f.write(f"}}\n\n")
        
        print(f"✓ Nginx redirects saved to: {output_file}")
    
    def print_summary(self):
        """Print a summary of mappings to console."""
        if not self.mappings:
            print("\nNo files were moved or renamed.")
            return
        
        print(f"\nSample Mappings (first 10):")
        print(f"{'='*70}")
        
        for i, mapping in enumerate(self.mappings[:10], 1):
            print(f"\n{i}. {mapping['filename']}")
            print(f"   OLD: {mapping['old_url']}")
            print(f"   NEW: {mapping['new_url']}")
        
        if len(self.mappings) > 10:
            print(f"\n... and {len(self.mappings) - 10} more")
        
        print(f"\n{'='*70}")


def main():
    """Main execution function."""
    
    print("\n" + "="*70)
    print("URL Mapping Generator for File Reorganization")
    print("="*70)
    
    # Get input files
    print("\nYou need two inventory files: BEFORE and AFTER the reorganization")
    print("These should be from the same directory (e.g., /uploaded-files/cropped-images)\n")
    
    before_file = input("BEFORE inventory file (JSON or CSV): ").strip()
    after_file = input("AFTER inventory file (JSON or CSV): ").strip()
    
    if not (Path(before_file).exists() and Path(after_file).exists()):
        print("✗ One or both inventory files not found")
        return
    
    # Get base URL
    base_url = input("\nWebsite base URL (e.g., https://example.com) [optional]: ").strip()
    
    # Generate mappings
    generator = URLMappingGenerator(base_url=base_url)
    mappings = generator.generate_mappings(before_file, after_file)
    
    if not mappings:
        print("\n✓ No files were moved or renamed. No mappings to generate.")
        return
    
    # Save in multiple formats
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    generator.save_csv(f"url_mappings_{timestamp}.csv")
    generator.save_json(f"url_mappings_{timestamp}.json")
    generator.save_htaccess(f"redirects_{timestamp}.htaccess")
    generator.save_nginx(f"redirects_{timestamp}.nginx.conf")
    
    # Print summary
    generator.print_summary()
    
    print(f"\n{'='*70}")
    print("✓ All mapping files generated successfully!")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    main()
