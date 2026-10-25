"""
AutoFlow Studio - File Classifier & Duplicate Remover
Safely categorizes files and detects duplicates using cryptographic hashing.
"""

import os
import hashlib
import shutil

CATEGORIES = {
    'Images': ['jpg', 'jpeg', 'png', 'gif', 'svg', 'webp', 'bmp'],
    'Documents': ['pdf', 'docx', 'doc', 'txt', 'rtf', 'odt', 'md'],
    'Datasets': ['csv', 'xlsx', 'xls', 'json', 'parquet', 'xml'],
    'Archives': ['zip', 'tar', 'gz', '7z', 'rar'],
    'Code': ['py', 'js', 'ts', 'html', 'css', 'json']
}

def get_file_checksum(filepath: str, chunk_size: int = 8192) -> str:
    """Computes MD5 hash of file content."""
    md5 = hashlib.md5()
    with open(filepath, 'rb') as f:
        while chunk := f.read(chunk_size):
            md5.update(chunk)
    return md5.hexdigest()

def organize_and_deduplicate(target_dir: str, dry_run: bool = True):
    if not os.path.exists(target_dir):
        raise FileNotFoundError(f"Target directory '{target_dir}' does not exist.")

    seen_hashes = {}
    duplicates = []
    actions = []

    # 1. Deduplicate check
    for item in os.listdir(target_dir):
        item_path = os.path.join(target_dir, item)
        if os.path.isfile(item_path):
            file_hash = get_file_checksum(item_path)
            if file_hash in seen_hashes:
                duplicates.append((item, seen_hashes[file_hash]))
                actions.append(f"[DUPLICATE] {item} matches {seen_hashes[file_hash]}")
            else:
                seen_hashes[file_hash] = item
                
                # 2. Categorize
                ext = item.split('.')[-1].lower() if '.' in item else 'other'
                category = 'Other'
                for cat_name, extensions in CATEGORIES.items():
                    if ext in extensions:
                        category = cat_name
                        break
                
                dest_dir = os.path.join(target_dir, category)
                actions.append(f"[MOVE] {item} -> {category}/")
                
                if not dry_run:
                    os.makedirs(dest_dir, exist_ok=True)
                    shutil.move(item_path, os.path.join(dest_dir, item))

    print(f"Summary for {target_dir}:")
    print(f"Total Unique Files: {len(seen_hashes)}")
    print(f"Duplicate Files: {len(duplicates)}")
    for act in actions:
        print(" ", act)

if __name__ == "__main__":
    organize_and_deduplicate(".", dry_run=True)
