#!/usr/bin/env python3
"""
Analyze which files in /media are unique vs duplicated elsewhere
"""

import csv
from collections import defaultdict

# Load the inventory
files_by_hash = defaultdict(list)
media_files = []

with open('ftp_inventory_20260120_174025.csv', 'r', encoding='utf-8') as f:
    reader = csv.DictReader(f)
    for row in reader:
        path = row['path']
        md5 = row['md5']
        
        # Track all files by hash
        if md5 and md5 != 'ERROR':
            files_by_hash[md5].append(path)
        
        # Track files in /media
        if path.startswith('/media/'):
            media_files.append(row)

# Find media files that are unique (no duplicates elsewhere)
unique_to_media = []
duplicated = []

for file_info in media_files:
    md5 = file_info['md5']
    path = file_info['path']
    
    if md5 == 'ERROR':
        print(f"WARNING - No hash: {path}")
        continue
    
    all_locations = files_by_hash[md5]
    
    # Check if this file exists outside /media
    non_media_locations = [loc for loc in all_locations if not loc.startswith('/media/')]
    
    if non_media_locations:
        duplicated.append({
            'media_path': path,
            'also_at': non_media_locations,
            'size': file_info['size']
        })
    else:
        unique_to_media.append(file_info)

# Report results
print(f"\n{'='*70}")
print(f"Analysis of /media folder")
print(f"{'='*70}\n")
print(f"Total files in /media: {len(media_files)}")
print(f"Files UNIQUE to /media (not found elsewhere): {len(unique_to_media)}")
print(f"Files duplicated outside /media (safe to delete): {len(duplicated)}\n")

if unique_to_media:
    print(f"{'='*70}")
    print(f"UNIQUE FILES - These exist ONLY in /media:")
    print(f"{'='*70}\n")
    for f in unique_to_media:
        print(f"{f['path']} ({f['size']} bytes)")
    print()

if duplicated:
    print(f"{'='*70}")
    print(f"DUPLICATED FILES - These exist elsewhere (first 20):")
    print(f"{'='*70}\n")
    for i, dup in enumerate(duplicated[:20], 1):
        print(f"{i}. {dup['media_path']}")
        print(f"   Also at: {dup['also_at'][0]}")
        if len(dup['also_at']) > 1:
            print(f"   (and {len(dup['also_at'])-1} other location(s))")
        print()
    
    if len(duplicated) > 20:
        print(f"... and {len(duplicated)-20} more duplicated files\n")
