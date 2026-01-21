#!/usr/bin/env python3
"""
FTP File Inventory Generator
Creates a comprehensive inventory of files from an FTP server with MD5 hashes
for tracking file relocations and URL mapping.
"""

import ftplib
import hashlib
import csv
import json
from pathlib import Path
from datetime import datetime
from typing import List, Dict
import os
from io import BytesIO


class FTPInventory:
    """Safely connect to FTP and generate file inventory with hashes."""
    
    def __init__(self, host: str, username: str, password: str, port: int = 21, timeout: int = 60):
        """
        Initialize FTP connection parameters.
        
        Args:
            host: FTP server hostname
            username: FTP username
            password: FTP password
            port: FTP port (default: 21)
            timeout: Connection timeout in seconds (default: 60)
        """
        self.host = host
        self.username = username
        self.password = password
        self.port = port
        self.timeout = timeout
        self.ftp = None
        self.inventory: List[Dict] = []
    
    def connect(self):
        """Establish FTP connection with error handling."""
        try:
            self.ftp = ftplib.FTP()
            self.ftp.connect(self.host, self.port, timeout=self.timeout)
            self.ftp.login(self.username, self.password)
            # Enable passive mode for better firewall compatibility
            self.ftp.set_pasv(True)
            print(f"✓ Connected to {self.host}")
            return True
        except ftplib.error_perm as e:
            print(f"✗ Permission error: {e}")
            return False
        except Exception as e:
            print(f"✗ Connection failed: {e}")
            return False
    
    def disconnect(self):
        """Safely close FTP connection."""
        if self.ftp:
            try:
                self.ftp.quit()
                print("✓ Disconnected from FTP server")
            except:
                self.ftp.close()
    
    def calculate_md5(self, filepath: str, max_retries: int = 3) -> str:
        """
        Download file from FTP and calculate MD5 hash with retry logic.
        
        Args:
            filepath: Path to file on FTP server
            max_retries: Number of retry attempts
            
        Returns:
            MD5 hash as hex string
        """
        for attempt in range(max_retries):
            md5_hash = hashlib.md5()
            buffer = BytesIO()
            
            try:
                # Send NOOP to keep connection alive
                try:
                    self.ftp.voidcmd('NOOP')
                except:
                    # Connection might be dead, reconnect
                    self.disconnect()
                    self.connect()
                
                # Download file to memory with timeout handling
                self.ftp.retrbinary(f'RETR {filepath}', buffer.write)
                buffer.seek(0)
                
                # Calculate hash in chunks
                for chunk in iter(lambda: buffer.read(8192), b''):
                    md5_hash.update(chunk)
                
                return md5_hash.hexdigest()
            except (ftplib.error_temp, EOFError, TimeoutError, OSError) as e:
                if attempt < max_retries - 1:
                    print(f"  Retry {attempt + 1}/{max_retries} (connection issue)")
                    # Reconnect and try again
                    try:
                        self.disconnect()
                        self.connect()
                    except:
                        pass
                else:
                    print(f"  Warning: Could not hash {filepath}: {e}")
                    return "ERROR"
            except Exception as e:
                print(f"  Warning: Could not hash {filepath}: {e}")
                return "ERROR"
        
        return "ERROR"
    
    def list_files_recursive(self, path: str = "/") -> List[Dict]:
        """
        Recursively list all files in FTP directory.
        
        Args:
            path: Starting path on FTP server
            
        Returns:
            List of file information dictionaries
        """
        files = []
        
        try:
            # Change to directory
            self.ftp.cwd(path)
            
            # Get directory listing
            items = []
            self.ftp.retrlines('LIST', items.append)
            
            for item in items:
                # Parse FTP LIST format (Unix-style)
                parts = item.split(None, 8)
                if len(parts) < 9:
                    continue
                
                permissions = parts[0]
                size = parts[4]
                name = parts[8]
                
                # Skip current and parent directory references
                if name in ['.', '..']:
                    continue
                
                full_path = f"{path}/{name}".replace('//', '/')
                
                # Check if it's a directory (starts with 'd')
                if permissions.startswith('d'):
                    # Recursively process subdirectory
                    print(f"  Scanning directory: {full_path}")
                    files.extend(self.list_files_recursive(full_path))
                else:
                    # It's a file
                    files.append({
                        'path': full_path,
                        'name': name,
                        'size': size,
                        'permissions': permissions
                    })
            
            # Return to parent directory
            self.ftp.cwd('..')
            
        except ftplib.error_perm as e:
            print(f"  Warning: Cannot access {path}: {e}")
        
        return files
    
    def generate_inventory(self, start_path: str = "/", include_hashes: bool = True, reconnect_interval: int = 50):
        """
        Generate complete inventory with optional hash calculation.
        
        Args:
            start_path: Starting directory path
            include_hashes: Whether to calculate MD5 hashes (slower but recommended)
            reconnect_interval: Reconnect to FTP every N files to prevent timeout (default: 50)
        """
        print(f"\n{'='*60}")
        print(f"Starting inventory generation from: {start_path}")
        print(f"{'='*60}\n")
        
        # Get file list
        print("Phase 1: Scanning directory structure...")
        files = self.list_files_recursive(start_path)
        print(f"✓ Found {len(files)} files\n")
        
        # Calculate hashes if requested
        if include_hashes:
            print("Phase 2: Calculating MD5 hashes...")
            print(f"(Reconnecting every {reconnect_interval} files to prevent timeout)\n")
            
            error_count = 0
            for i, file_info in enumerate(files, 1):
                # Reconnect periodically to prevent timeout
                if i % reconnect_interval == 0:
                    print(f"  [Progress: {i}/{len(files)} - Reconnecting...]")
                    try:
                        self.disconnect()
                        self.connect()
                    except Exception as e:
                        print(f"  Warning: Reconnection failed: {e}")
                
                print(f"  [{i}/{len(files)}] {file_info['path']}")
                file_info['md5'] = self.calculate_md5(file_info['path'])
                file_info['timestamp'] = datetime.now().isoformat()
                
                if file_info['md5'] == 'ERROR':
                    error_count += 1
            
            print()
            if error_count > 0:
                print(f"⚠ Warning: {error_count} files could not be hashed")
                print(f"  You may want to run the script again to retry these files.\n")
        
        self.inventory = files
        return files
    
    def save_to_csv(self, output_file: str = "ftp_inventory.csv"):
        """Save inventory to CSV file."""
        if not self.inventory:
            print("No inventory data to save")
            return
        
        with open(output_file, 'w', newline='', encoding='utf-8') as f:
            if self.inventory:
                fieldnames = self.inventory[0].keys()
                writer = csv.DictWriter(f, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(self.inventory)
        
        print(f"✓ CSV saved to: {output_file}")
    
    def save_to_json(self, output_file: str = "ftp_inventory.json"):
        """Save inventory to JSON file."""
        if not self.inventory:
            print("No inventory data to save")
            return
        
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump({
                'generated_at': datetime.now().isoformat(),
                'total_files': len(self.inventory),
                'files': self.inventory
            }, f, indent=2)
        
        print(f"✓ JSON saved to: {output_file}")


def main():
    """Main execution function with safe credential handling."""
    
    # Get credentials from environment variables (recommended)
    # Set these in your shell: export FTP_HOST=ftp.example.com
    host = os.getenv('FTP_HOST', '')
    username = os.getenv('FTP_USER', '')
    password = os.getenv('FTP_PASS', '')
    
    # If not in environment, prompt user
    if not all([host, username, password]):
        print("FTP Credentials not found in environment variables.")
        print("Please enter credentials (or press Ctrl+C to cancel):\n")
        host = input("FTP Host: ").strip()
        username = input("Username: ").strip()
        password = input("Password: ").strip()  # Note: visible on screen
    
    if not all([host, username, password]):
        print("✗ Missing required credentials")
        return
    
    # Optional: specify starting directory
    start_path = input("Starting directory path (default '/', e.g., '/uploaded-files/cropped-images'): ").strip() or "/"
    
    # Create inventory instance
    inventory = FTPInventory(host, username, password)
    
    try:
        # Connect to FTP
        if not inventory.connect():
            return
        
        # Generate inventory with hashes
        inventory.generate_inventory(start_path=start_path, include_hashes=True)
        
        # Save results
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        inventory.save_to_csv(f"ftp_inventory_{timestamp}.csv")
        inventory.save_to_json(f"ftp_inventory_{timestamp}.json")
        
        print(f"\n{'='*60}")
        print(f"✓ Inventory complete! Total files: {len(inventory.inventory)}")
        print(f"{'='*60}\n")
        
    finally:
        # Always disconnect
        inventory.disconnect()


if __name__ == "__main__":
    main()
