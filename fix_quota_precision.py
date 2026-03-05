#!/usr/bin/env python3
"""
Script to fix corrupted 定额 values in the quota table.
The corrupted values are floating-point precision errors that occurred during data import.
These values are extremely large (e.g., 25155000000000004) when they should be much smaller.
"""

import sqlite3
import os
from decimal import Decimal, ROUND_HALF_UP

def get_db_path():
    """Get database path from environment variable or use default."""
    return os.environ.get('SQLITE_DB_PATH', 'payroll_database.db')

def identify_corrupted_values(cursor):
    """
    Identify records with likely corrupted 定额 values.
    Returns a list of tuples (code, process, current_value).
    """
    # Values that are extremely large (over 1,000,000) are likely corrupted
    # Also check for values ending with common floating-point error patterns
    cursor.execute('''
        SELECT 代码, 加工工序, 定额 
        FROM quota 
        WHERE 定额 > 1000000
           OR 定额 LIKE '%00000000000004'
           OR 定额 LIKE '%99999999999996'
           OR 定额 LIKE '%99999999999998'
           OR 定额 LIKE '%00000000000002'
        ORDER BY CAST(定额 AS REAL) DESC
    ''')
    return cursor.fetchall()

def try_correct_value(value_str):
    """
    Try to correct a corrupted floating-point value.
    This function attempts to divide by powers of 10 to find a reasonable value.
    Returns the corrected value if successful, or None if correction is not possible.
    """
    try:
        value = float(value_str)
        
        # If value is already reasonable (< 1000000), return as is
        if value < 1000000:
            return round(value, 2)
        
        # Try dividing by 10^10, 10^11, 10^12, 10^13, 10^14 to find a reasonable value
        for power in range(10, 15):
            divisor = 10 ** power
            corrected = value / divisor
            if 0.01 < corrected < 1000000:
                return round(corrected, 2)
        
        # If still too large, try rounding to 2 decimal places
        return round(value, 2)
    except (ValueError, TypeError):
        return None

def fix_quota定额_precision(dry_run=True, fix_all=False):
    """
    Fix corrupted 定额 values in the quota table.
    
    Args:
        dry_run: If True, only show what would be changed without making changes.
        fix_all: If True, fix all corrupted values automatically.
    """
    db_path = get_db_path()
    print(f'Using database: {db_path}')
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Get corrupted values
    corrupted = identify_corrupted_values(cursor)
    total_corrupted = len(corrupted)
    
    print(f'\nFound {total_corrupted} records with potentially corrupted 定额 values')
    
    if total_corrupted == 0:
        print('No corrupted values found.')
        conn.close()
        return
    
    # Show sample of corrupted values
    print('\nSample of corrupted values:')
    for row in corrupted[:10]:
        corrected = try_correct_value(str(row[2]))
        print(f'  代码={row[0]}, 工序={row[1]}')
        print(f'    Current:  {row[2]}')
        print(f'    Corrected: {corrected}')
        print()
    
    if total_corrupted > 10:
        print(f'... and {total_corrupted - 10} more records')
    
    if dry_run:
        print('\n=== DRY RUN - No changes made ===')
        print(f'Would fix {total_corrupted} records.')
        conn.close()
        return
    
    if not fix_all:
        print('\nDo you want to fix these values? (y/n/all): ', end='')
        choice = input().strip().lower()
        
        if choice not in ['y', 'yes', 'all']:
            print('Aborted.')
            conn.close()
            return
    else:
        choice = 'all'
    
    # Fix the corrupted values
    fixed_count = 0
    error_count = 0
    
    for code, process, current_value in corrupted:
        corrected = try_correct_value(str(current_value))
        
        if corrected is not None:
            try:
                cursor.execute('''
                    UPDATE quota 
                    SET 定额 = ? 
                    WHERE 代码 = ? AND 加工工序 = ?
                ''', (corrected, code, process))
                fixed_count += 1
            except Exception as e:
                print(f'Error updating 代码={code}, 工序={process}: {e}')
                error_count += 1
        else:
            error_count += 1
    
    conn.commit()
    conn.close()
    
    print(f'\nFixed {fixed_count} records')
    if error_count > 0:
        print(f'Errors: {error_count} records')

def show_quota_stats():
    """Show statistics about the quota table."""
    db_path = get_db_path()
    
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    
    # Total records
    cursor.execute('SELECT COUNT(*) FROM quota')
    total = cursor.fetchone()[0]
    
    # Records with 定额 > 1000000
    cursor.execute('SELECT COUNT(*) FROM quota WHERE 定额 > 1000000')
    corrupted = cursor.fetchone()[0]
    
    # Min, max, avg 定额
    cursor.execute('SELECT MIN(定额), MAX(定额), AVG(定额) FROM quota')
    min_val, max_val, avg_val = cursor.fetchone()
    
    # Records with likely floating-point errors
    cursor.execute('''
        SELECT COUNT(*) FROM quota 
        WHERE 定额 LIKE '%00000000000004'
           OR 定额 LIKE '%99999999999996'
           OR 定额 LIKE '%99999999999998'
           OR 定额 LIKE '%00000000000002'
    ''')
    fp_errors = cursor.fetchone()[0]
    
    conn.close()
    
    print('\n=== Quota Table Statistics ===')
    print(f'Total records: {total}')
    print(f'Records with 定额 > 1,000,000: {corrupted}')
    print(f'Records with likely floating-point errors: {fp_errors}')
    print(f'Min 定额: {min_val}')
    print(f'Max 定额: {max_val}')
    print(f'Avg 定额: {avg_val:.2f}')

if __name__ == '__main__':
    import sys
    
    dry_run = '--dry-run' in sys.argv or '-n' in sys.argv
    fix_all = '--fix-all' in sys.argv or '-f' in sys.argv
    stats = '--stats' in sys.argv or '-s' in sys.argv
    
    if stats:
        show_quota_stats()
    else:
        fix_quota定额_precision(dry_run=dry_run, fix_all=fix_all)
