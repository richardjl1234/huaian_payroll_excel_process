#!/usr/bin/env python3
"""
Script to update the SQLite database schema from FLOAT to NUMERIC(10,2)
for monetary fields to ensure precision.
"""

import sqlite3
import os
from excel_processor.config import DATABASE_PATH

def update_database_schema():
    """
    Update the database schema to change FLOAT columns to NUMERIC(10,2)
    """
    try:
        # Connect to the database
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        
        print("Connected to database successfully")
        
        # Check if payroll_details table exists
        cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='payroll_details';")
        table_exists = cursor.fetchone()
        
        if not table_exists:
            print("Error: payroll_details table does not exist")
            return False
        
        # Get current table schema
        cursor.execute("PRAGMA table_info(payroll_details);")
        columns = cursor.fetchall()
        
        print("Current table schema:")
        for col in columns:
            print(f"  {col[1]}: {col[2]}")
        
        # Create a temporary table with the new schema
        create_temp_table_sql = """
        CREATE TABLE IF NOT EXISTS payroll_details_temp (
            文件名 CHAR(100),
            sheet名 CHAR(100),
            职员全名 CHAR(20),
            日期 CHAR(20),
            客户名称 CHAR(60),
            型号 CHAR(100),
            工序全名 CHAR(100),
            工序 CHAR(100),
            计件数量 NUMERIC(10,2),
            系数 NUMERIC(10,2),
            定额 NUMERIC(10,2),
            金额 NUMERIC(10,2),
            备注 CHAR(100)
        )
        """
        cursor.execute(create_temp_table_sql)
        
        # Copy data from old table to new table
        copy_data_sql = """
        INSERT INTO payroll_details_temp 
        SELECT 
            文件名, sheet名, 职员全名, 日期, 客户名称, 型号, 工序全名, 工序,
            CAST(计件数量 AS NUMERIC(10,2)),
            CAST(系数 AS NUMERIC(10,2)),
            CAST(定额 AS NUMERIC(10,2)),
            CAST(金额 AS NUMERIC(10,2)),
            备注
        FROM payroll_details
        """
        cursor.execute(copy_data_sql)
        
        # Drop the old table
        cursor.execute("DROP TABLE payroll_details")
        
        # Rename the temporary table to the original name
        cursor.execute("ALTER TABLE payroll_details_temp RENAME TO payroll_details")
        
        # Commit changes
        conn.commit()
        
        # Verify the new schema
        cursor.execute("PRAGMA table_info(payroll_details);")
        new_columns = cursor.fetchall()
        
        print("\nNew table schema:")
        for col in new_columns:
            print(f"  {col[1]}: {col[2]}")
        
        # Count records to verify data integrity
        cursor.execute("SELECT COUNT(*) FROM payroll_details")
        record_count = cursor.fetchone()[0]
        print(f"\nData migration completed successfully. Total records: {record_count}")
        
        conn.close()
        return True
        
    except Exception as e:
        print(f"Error updating database schema: {e}")
        if conn:
            conn.rollback()
            conn.close()
        return False

if __name__ == "__main__":
    print("Starting database schema update...")
    success = update_database_schema()
    if success:
        print("Database schema update completed successfully!")
    else:
        print("Database schema update failed!")
