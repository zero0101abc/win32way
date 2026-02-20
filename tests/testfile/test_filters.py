#!/usr/bin/env python3
"""
Test script for EmailFilterManager
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from test import EmailFilterManager

def main():
    print("=== Testing EmailFilterManager ===")
    
    # Initialize filter manager
    filter_manager = EmailFilterManager()
    
    # Don't clear existing filters - add to them instead
    
    print("\n1. Creating MX filter with null settings...")
    # Create MX filter with null/empty settings
    mx_success = filter_manager.create_filter(
        name="mx",
        from_email="",
        subject_filter="",
        body_filter="",
        action="",
        description="MX filter - all settings null initially"
    )
    print(f"MX filter created: {mx_success}")
    
    print("\n2. Creating CDC filter with null settings...")
    # Create CDC filter with null/empty settings
    cdc_success = filter_manager.create_filter(
        name="cdc", 
        from_email="",
        subject_filter="",
        body_filter="",
        action="",
        description="CDC filter - all settings null initially"
    )
    print(f"CDC filter created: {cdc_success}")
    
    print("\n3. Listing all filters...")
    filters = filter_manager.list_filters()
    for f in filters:
        print(f"ID: {f['id']}, Name: {f['name']}, From: '{f['from_email']}', Subject: '{f['subject_filter']}', Action: '{f['action']}'")
    
    print("\n4. Testing filter editing...")
    # Edit MX filter
    mx_edit_success = filter_manager.edit_filter(
        1,  # MX filter ID
        from_email="system.MX@hkt-emsconnect.com",
        subject_filter='contains(subject, "MX Alert")',
        action="send_mx_alert"
    )
    print(f"MX filter edited: {mx_edit_success}")
    
    # Edit CDC filter
    cdc_edit_success = filter_manager.edit_filter(
        2,  # CDC filter ID
        from_email="cdc@notifications.com",
        subject_filter='equals(subject, "CDC Report")',
        body_filter='contains(body, "statistics")',
        action="log_cdc_data"
    )
    print(f"CDC filter edited: {cdc_edit_success}")
    
    print("\n5. Updated filters:")
    filters = filter_manager.list_filters()
    for f in filters:
        print(f"ID: {f['id']}, Name: {f['name']}")
        print(f"  From: {f['from_email']}")
        print(f"  Subject: {f['subject_filter']}")
        print(f"  Body: {f['body_filter']}")
        print(f"  Action: {f['action']}")
        print(f"  Description: {f['description']}")
        print()
    
    print("6. Testing with sample email data...")
    # Test with sample email
    sample_email = {
        "sender": "system.MX@hkt-emsconnect.com",
        "subject": "MX Alert - System Warning",
        "body": "This is an MX alert with statistics",
        "date": "2026-01-26",
        "recipients": []
    }
    
    actions = filter_manager.apply_filters(sample_email)
    print(f"Sample email actions: {actions}")
    
    print("\n=== Test Complete ===")
    print("Filters saved to: email_filters.json")

if __name__ == "__main__":
    main()