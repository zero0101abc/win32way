#!/usr/bin/env python3

import json
import sys
import os

# Add the current directory to Python path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from test import EmailFilterManager

def test_cdc_filter():
    """Test CDC filter with sample email data"""
    
    # Load the email filter manager
    filter_manager = EmailFilterManager()
    
    # Sample CDC email from outlook_emails.json
    sample_email = {
        "sender": "CDC ITD CallCenter",
        "date": "2026-01-21 11:30:20.134000+00:00", 
        "subject": "[HK540639]灣仔中信(601)，IT - System / Software，Digital Menu，Notebook - PPN OS，notebook無法更新餐倉的餐條，",
        "body": """CAUTION: This email originated from outside of the organisation. Do not click links or open attachments unless you recognise the sender and know the content is safe. 
警告： 這封電子郵件來自外來電郵。如不認識寄件者，或不確定內容是否安全，切勿點開其中的連結，或是打開或儲存附件。 
Dear ,CDC: Vendor- PPN
ITD-Ticket# (HK540639) has been reported on [2026-01-21 11:29:32] by [彭小姐].
[CDC]/[灣仔中信(601)]

Inci. ID:

HK540639 

Cust. Name:

灣仔中信(601) 

Type:

IT - System / Software 

Category:

Digital Menu 

Sub-Category:

Notebook - PPN OS 

Priority:

High 

Handling Office:

CDC: Vendor- PPN 

Reporter Name:

彭小姐 

Contact Number 1:

66215520 

Contact email:

cdc.601@cafedecoral.com, 

Description:

notebook無法更新餐倉的餐條 


Please refer to the Inci. ID for all future correspondence about this issue.

Best Regards，
ITD Call Center""",
        "recipients": [
            {
                "name": "iSupport",
                "email": "/o=ExchangeLabs/ou=Exchange Administrative Group (FYDIBOHF23SPDLT)/cn=Recipients/cn=1fe8c29757374bfeae9dbafd2187f2bb-20169a46-cf",
                "type": 1
            },
            {
                "name": "operator@cafedecoral.com", 
                "email": "operator@cafedecoral.com",
                "type": 2
            },
            {
                "name": "cdc_hk_itd@wtto.com",
                "email": "cdc_hk_itd@wtto.com",
                "type": 2
            }
        ]
    }
    
    print("Testing CDC filter with sample email...")
    print(f"Sender: {sample_email['sender']}")
    print(f"Subject: {sample_email['subject'][:50]}...")
    print(f"Recipients: {[r['email'] for r in sample_email['recipients']]}")
    print()
    
    # List all filters
    print("Available filters:")
    for filter_data in filter_manager.filters:
        print(f"  ID: {filter_data['id']}, Name: {filter_data['name']}")
        print(f"    From: {filter_data['from_email']}")
        print(f"    Subject: {filter_data['subject_filter']}")
        print(f"    To: {filter_data['to_email']}")
        print(f"    Enabled: {filter_data['enabled']}")
        print()
    
    # Test each filter step by step
    print("Testing each filter step by step:")
    for filter_data in filter_manager.filters:
        if not filter_data.get('enabled', True):
            print(f"Filter {filter_data['name']} is disabled - skipping")
            continue
            
        print(f"\nTesting filter: {filter_data['name']} (ID: {filter_data['id']})")
        
        # Check from email
        from_match = True
        if filter_data.get('from_email'):
            from_match = filter_data['from_email'] in sample_email.get('sender', '')
            print(f"  From email check: {filter_data['from_email']} in {sample_email['sender']} = {from_match}")
        
        # Check subject filter
        subject_match = True
        if filter_data.get('subject_filter'):
            try:
                subject_match = filter_manager.evaluate_power_automate_expression(
                    filter_data['subject_filter'], sample_email)
                print(f"  Subject filter check: {filter_data['subject_filter']} = {subject_match}")
            except Exception as e:
                print(f"  Subject filter ERROR: {e}")
                subject_match = False
        
        # Check to email
        to_match = True
        if filter_data.get('to_email'):
            to_recipients = [r['email'] for r in sample_email.get('recipients', []) if r.get('type') == 1]
            to_match = any(filter_data['to_email'] in to_email for to_email in to_recipients)
            print(f"  To email check: {filter_data['to_email']} in {to_recipients} = {to_match}")
        
        overall_match = from_match and subject_match and to_match
        print(f"  Overall match: {overall_match}")
        
        if overall_match and filter_data.get('action'):
            print(f"  ✓ FILTER MATCHED! Action: {filter_data['action']}")
    
    # Test apply_filters method
    print("\n" + "="*50)
    print("Testing apply_filters method:")
    try:
        actions = filter_manager.apply_filters(sample_email)
        print(f"Returned actions: {actions}")
        print(f"Actions length: {len(actions)}")
    except Exception as e:
        print(f"Error in apply_filters: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_cdc_filter()
