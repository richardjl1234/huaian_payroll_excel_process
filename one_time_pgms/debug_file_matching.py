#!/usr/bin/env python3
"""
Debug script to check file matching between old and new folders
"""

import os
import glob

# Paths to the folders
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OLD_FOLDER = os.path.join(BASE_DIR, 'old_payroll')
NEW_FOLDER = os.path.join(BASE_DIR, 'new_payroll')

def debug_file_matching():
    """Debug file matching logic"""
    print("=== DEBUG FILE MATCHING ===")
    
    # Get files from old folder
    old_files = []
    for ext in ['*.xls', '*.xlsx']:
        old_files.extend(glob.glob(os.path.join(OLD_FOLDER, ext)))
    
    # Get files from new folder
    new_files = []
    for ext in ['*.xls', '*.xlsx']:
        new_files.extend(glob.glob(os.path.join(NEW_FOLDER, ext)))
    
    print(f"Old folder files found: {len(old_files)}")
    print(f"New folder files found: {len(new_files)}")
    
    # Extract base names without extension
    old_basenames = {os.path.splitext(os.path.basename(f))[0]: f for f in old_files}
    new_basenames = {os.path.splitext(os.path.basename(f))[0]: f for f in new_files}
    
    print(f"\nOld folder basenames: {sorted(old_basenames.keys())}")
    print(f"\nNew folder basenames: {sorted(new_basenames.keys())}")
    
    # Find common files that exist in both folders
    common_files = sorted(set(old_basenames.keys()) & set(new_basenames.keys()))
    
    print(f"\nCommon files: {common_files}")
    
    # Check specifically for 202001-202012 files
    target_files = [f"2020{str(i).zfill(2)}" for i in range(1, 13)]
    print(f"\nTarget files (202001-202012): {target_files}")
    
    available_targets = []
    for target in target_files:
        old_has = target in old_basenames
        new_has = target in new_basenames
        status = "BOTH" if old_has and new_has else "OLD_ONLY" if old_has else "NEW_ONLY" if new_has else "NONE"
        available_targets.append((target, status))
        print(f"  {target}: {status}")
    
    return common_files

if __name__ == '__main__':
    debug_file_matching()
