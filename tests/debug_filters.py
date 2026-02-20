#!/usr/bin/env python3
import json
import sys

# Set UTF-8 encoding for output
if sys.platform == "win32":
    import codecs
    sys.stdout = codecs.getwriter("utf-8")(sys.stdout.detach())

def contains(text: str, search: str) -> bool:
    return search.lower() in text.lower()

def safe_str(text):
    try:
        return str(text)
    except:
        return "[UNICODE_ERROR]"

def debug_filtering():
    # Load data
    with open('outlook_emails.json', 'r', encoding='utf-8') as f:
        emails = json.load(f)
    
    with open('email_filters.json', 'r', encoding='utf-8') as f:
        filters = json.load(f)
    
    print("="*80)
    print("DEBUG FILTERING ANALYSIS")
    print("="*80)
    
    cdc_count = 0
    mx_count = 0
    
    for i, email in enumerate(emails[:5]):  # Just check first 5 emails
        print(f"\n[Email {i+1}]")
        sender = safe_str(email.get('sender', 'N/A'))
        subject = safe_str(email.get('subject', 'N/A'))
        print(f"  Sender: {sender}")
        print(f"  Subject: {subject[:50]}...")
        
        # Show recipients
        recipients = email.get('recipients', [])
        to_names = [safe_str(r['name']) for r in recipients if r.get('type') == 1]
        cc_names = [safe_str(r['name']) for r in recipients if r.get('type') == 2]
        print(f"  TO: {to_names}")
        print(f"  CC: {cc_names}")
        
        # Test each filter
        for filter_data in filters:
            if not filter_data.get('enabled', True):
                continue
            
            filter_name = filter_data.get('name', 'Unknown')
            print(f"\n  Testing filter: {filter_name}")
            
            # Check from email filter
            from_email_match = True
            if filter_data.get('from_email'):
                sender_matches = contains(email.get('sender', ''), filter_data['from_email'])
                print(f"    From email check: '{filter_data['from_email']}' in '{sender}' = {sender_matches}")
                from_email_match = sender_matches
                
                if not sender_matches:
                    print(f"    -> FILTER FAILED: From email doesn't match")
                    continue
            
            # Check subject filter 
            subject_match = True
            if filter_data.get('subject_filter'):
                if 'has been assigned to' in filter_data['subject_filter']:
                    subject_has = 'has been assigned to' in email.get('subject', '')
                    print(f"    Subject check: 'has been assigned to' in subject = {subject_has}")
                    subject_match = subject_has
                    
                    if not subject_has:
                        print(f"    -> FILTER FAILED: Subject doesn't contain 'has been assigned to'")
                        continue
            
            # Check body filter
            if filter_data.get('body_filter'):
                print(f"    Body filter: {filter_data['body_filter']} (NOT IMPLEMENTED)")
                pass
            
            # Check TO email filter
            to_email_match = True
            if filter_data.get('to_email'):
                to_recipients = [r['name'] for r in email.get('recipients', []) if r.get('type') == 1]
                to_found = any(filter_data['to_email'] in to_name for to_name in to_recipients)
                print(f"    TO email check: '{filter_data['to_email']}' in {to_recipients} = {to_found}")
                to_email_match = to_found
                
                if not to_found:
                    print(f"    -> FILTER FAILED: TO email doesn't match")
                    continue
            
            # If all conditions match
            if from_email_match and subject_match and to_email_match:
                action = filter_data.get('action')
                print(f"    -> FILTER PASSED! Action: {action}")
                
                if action == 'extract_cdc':
                    cdc_count += 1
                elif action == 'send_mx_alert':
                    mx_count += 1
            else:
                print(f"    -> FILTER FAILED")
    
    print("\n" + "="*80)
    print("SUMMARY")
    print("="*80)
    print(f"CDC emails matched: {cdc_count}")
    print(f"MX emails matched: {mx_count}")
    print(f"Total emails processed: {min(5, len(emails))}")

if __name__ == "__main__":
    debug_filtering()
