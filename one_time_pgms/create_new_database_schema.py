#!/usr/bin/env python3
"""
Script to create a new database schema with 代码 column in payroll_details table
with foreign key constraint to quota table.
"""

import sqlite3
import os
from excel_processor.config import DATABASE_PATH

def create_new_database_schema():
    """
    Create a new database schema with the updated structure.
    This will drop existing tables and recreate them with the new schema.
    """
    try:
        # Connect to the database
        conn = sqlite3.connect(DATABASE_PATH)
        cursor = conn.cursor()
        
        print("Connected to database successfully")
        
        # Enable foreign key constraints
        cursor.execute("PRAGMA foreign_keys = ON;")
        
        # Drop existing tables if they exist
        print("Dropping existing tables...")
        cursor.execute("DROP TABLE IF EXISTS payroll_details;")
        cursor.execute("DROP TABLE IF EXISTS load_log;")
        
        # Create payroll_details table with 代码 column and foreign key constraint
        print("Creating payroll_details table with 代码 column...")
        create_payroll_details_sql = """
        CREATE TABLE payroll_details (
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
            备注 CHAR(100),
            代码 CHAR(12),
            FOREIGN KEY (代码) REFERENCES quota(代码)
        )
        """
        cursor.execute(create_payroll_details_sql)
        
        # Create load_log table
        print("Creating load_log table...")
        create_load_log_sql = """
        CREATE TABLE load_log (
            file_name CHAR(50),
            sheet_name CHAR(50),
            table_index INT,
            discarded_columns CHAR(200),
            discarded_cols_num INT
        )
        """
        cursor.execute(create_load_log_sql)
        
        # Commit changes
        conn.commit()
        
        # Verify the new schema
        print("\nNew payroll_details table schema:")
        cursor.execute("PRAGMA table_info(payroll_details);")
        for col in cursor.fetchall():
            print(f"  - {col[1]}: {col[2]}")
        
        # Check foreign key constraints
        cursor.execute("PRAGMA foreign_key_list(payroll_details);")
        foreign_keys = cursor.fetchall()
        if foreign_keys:
            print("\nForeign key constraints:")
            for fk in foreign_keys:
                print(f"  - {fk[3]} -> {fk[2]}.{fk[4]}")
        else:
            print("\nNo foreign key constraints found")
        
        conn.close()
        print("\nNew database schema created successfully!")
        return True
        
    except Exception as e:
        print(f"Error creating new database schema: {e}")
        if conn:
            conn.rollback()
            conn.close()
        return False

if __name__ == "__main__":
    print("Creating new database schema...")
    success = create_new_database_schema()
    if success:
        print("Database schema creation completed successfully!")
    else:
        print("Database schema creation failed!")
