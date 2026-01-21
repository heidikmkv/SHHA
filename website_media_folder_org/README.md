# Website Media Folder Organization

FTP-based file inventory system for tracking and reorganizing website media files.

## Features

- ✅ Secure FTP connection handling
- ✅ Recursive directory scanning
- ✅ MD5 hash calculation for each file
- ✅ CSV and JSON export formats
- ✅ Safe credential management via environment variables
- ✅ Progress tracking during scan

## Setup

1. **Install Python** (if needed):
   ```bash
   python3 --version  # Check if installed
   ```

2. **No dependencies needed** - uses only Python standard library!

3. **Configure credentials** (choose one method):

   **Option A: Environment Variables (Recommended)**
   ```bash
   export FTP_HOST="ftp.yoursite.com"
   export FTP_USER="your_username"
   export FTP_PASS="your_password"
   ```

   **Option B: .env file**
   ```bash
   cp .env.example .env
   # Edit .env with your credentials
   ```

   **Option C: Interactive prompt**
   - Just run the script, it will ask for credentials

## Usage

### Basic Usage
```bash
python3 ftp_inventory.py
```

### What It Does

1. Connects to your FTP server
2. Recursively scans all files in the specified directory
3. Downloads each file (to memory only) and calculates MD5 hash
4. Generates timestamped CSV and JSON files with:
   - Full file path
   - Filename
   - File size
   - MD5 hash
   - Timestamp

### Output Files

- `ftp_inventory_YYYYMMDD_HHMMSS.csv` - Spreadsheet format
- `ftp_inventory_YYYYMMDD_HHMMSS.json` - Structured data format

### Example Output (CSV)

```csv
path,name,size,permissions,md5,timestamp
/media/images/logo.png,logo.png,15234,-rw-r--r--,5d41402abc4b2a76b9719d911017c592,2026-01-20T10:30:00
/media/images/banner.jpg,banner.jpg,45678,-rw-r--r--,098f6bcd4621d373cade4e832627b4f6,2026-01-20T10:30:01
```

## URL Mapping

Once you have the inventory, you can:

1. Open the CSV in Excel/Google Sheets
2. Add a "new_url" column
3. Use the MD5 hash to identify duplicate files
4. Create redirects from old paths to new paths

## Security Notes

⚠️ **Important:**
- Never commit `.env` files or credentials to git
- The `.gitignore` is configured to protect credential files
- Consider using SSH/SFTP instead of FTP if available
- Files are downloaded to memory only (not saved to disk)

## Troubleshooting

**Connection Issues:**
- Verify FTP host, username, and password
- Check if FTP port is 21 (default) or custom
- Ensure firewall allows FTP connections

**Slow Performance:**
- Set `include_hashes=False` in code to skip MD5 calculation
- Process subdirectories separately

**Permission Errors:**
- Verify FTP user has read access to target directories
- Check that directory paths are correct

## Next Steps

After generating the inventory:

1. **Find Duplicates:**
   ```bash
   # Files with same MD5 are identical
   sort -t, -k5 ftp_inventory_*.csv | uniq -f4 -d
   ```

2. **Plan URL Structure:**
   - Organize by type (images, documents, videos)
   - Use meaningful directory names
   - Keep URLs short and SEO-friendly

3. **Generate Redirects:**
   - Create a mapping file from old → new URLs
   - Generate .htaccess or nginx redirect rules
