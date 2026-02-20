#!/usr/bin/env python3
"""
Test filter ID generation to ensure no duplicates
"""

import json
import os
from test import EmailFilterManager

def test_filter_id_generation():
    """Test that filter IDs are generated correctly without duplicates"""
    
    # Backup existing filters
    backup_file = "email_filters_backup.json"
    if os.path.exists("email_filters.json"):
        with open("email_filters.json", "r", encoding="utf-8") as f:
            backup_data = f.read()
        with open(backup_file, "w", encoding="utf-8") as f:
            f.write(backup_data)
    
    try:
        # Create fresh filter manager
        fm = EmailFilterManager()
        
        # Clear existing filters for clean test
        fm.filters = []
        fm.save_filters()
        
        print("Testing filter ID generation...")
        
        # Test 1: Create first filter
        result1 = fm.create_filter("Test Filter 1", "test1@example.com")
        print(f"Created filter 1: {result1}")
        filters = fm.list_filters()
        print(f"Filters after creation 1: {[f['id'] for f in filters]}")
        
        # Test 2: Create second filter
        result2 = fm.create_filter("Test Filter 2", "test2@example.com")
        print(f"Created filter 2: {result2}")
        filters = fm.list_filters()
        print(f"Filters after creation 2: {[f['id'] for f in filters]}")
        
        # Test 3: Create third filter
        result3 = fm.create_filter("Test Filter 3", "test3@example.com")
        print(f"Created filter 3: {result3}")
        filters = fm.list_filters()
        print(f"Filters after creation 3: {[f['id'] for f in filters]}")
        
        # Test 4: Delete filter 2
        delete_result = fm.delete_filter(2)
        print(f"Deleted filter 2: {delete_result}")
        filters = fm.list_filters()
        print(f"Filters after deletion: {[f['id'] for f in filters]}")
        
        # Test 5: Create new filter - should get ID 4, not 2
        result4 = fm.create_filter("Test Filter 4", "test4@example.com")
        print(f"Created filter 4: {result4}")
        filters = fm.list_filters()
        print(f"Filters after creation 4: {[f['id'] for f in filters]}")
        
        # Verify no duplicate IDs
        ids = [f['id'] for f in filters]
        unique_ids = set(ids)
        
        print(f"\nFinal filter IDs: {ids}")
        print(f"Unique IDs: {list(unique_ids)}")
        print(f"No duplicates: {len(ids) == len(unique_ids)}")
        
        # Expected result: IDs should be [1, 3, 4]
        expected_ids = [1, 3, 4]
        print(f"Expected IDs: {expected_ids}")
        print(f"IDs match expected: {ids == expected_ids}")
        
        return len(ids) == len(unique_ids) and ids == expected_ids
        
    finally:
        # Restore backup if it exists
        if os.path.exists(backup_file):
            with open(backup_file, "r", encoding="utf-8") as f:
                backup_data = f.read()
            with open("email_filters.json", "w", encoding="utf-8") as f:
                f.write(backup_data)
            os.remove(backup_file)

if __name__ == "__main__":
    success = test_filter_id_generation()
    print(f"\nTest {'PASSED' if success else 'FAILED'}")