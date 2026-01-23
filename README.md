# Data analysis & useful scripts for SHHA

This repository contains tools and data for managing the Seattle Horse History Archive (SHHA) digital collection.

## Projects

### GRIT_archive
Historical archive of GRIT (Grassroots Riders in Training) publications spanning 1979-2024.

**Contents:**
- **GRIT_archive_OCRtext/** - OCR-extracted text from 350+ GRIT newsletter issues for full-text search and analysis
- **data/** - Organized by year, containing OCR text output and metadata
- Thumbnails and AI-generated summaries (not tracked in git due to size)

### website_media_folder_org
FTP-based file inventory and URL mapping system for website media reorganization.

**Purpose:** 
Generate file inventories with MD5 hashes to safely track and redirect files during website restructuring.

**Tools:**
- `ftp_inventory.py` - Scan FTP directories and generate comprehensive file inventories with MD5 hashes
- `generate_url_mappings.py` - Compare before/after inventories to create URL redirect mappings
- Outputs redirect rules for Apache (.htaccess) and Nginx

**Workflow:**
1. Generate "before" inventory of current file structure
2. Reorganize files on FTP server
3. Generate "after" inventory
4. Automatically create old→new URL mappings for webmaster

### thumbnail_generator
Python script for generating thumbnail images from full-page scans.

### user_lists_analysis  
Jupyter notebook for analyzing SHHA user list data and engagement patterns.

## Data Organization

The repository follows a structure optimized for both human browsing and programmatic access, with large binary files excluded via `.gitignore` to keep the repository lightweight.

