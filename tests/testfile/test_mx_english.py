#!/usr/bin/env python3
"""
Test MX data extraction with sample English data
"""

import json
from test import EmailFilterManager

def test_mx_extraction_english():
    """Test the MX data extraction with English sample data"""
    sample_email = {
        "sender": "system.MX@hkt-emsconnect.com",
        "subject": "Incident has been assigned to your group",
        "body": """Incident Details:
Number: INC0012345
User: John Smith
Location: 0456-HK Central Office
Category: Hardware Issue
description: Server is down and needs immediate attention You can contact support team for assistance.
Additional notes: This is a critical issue affecting production systems.
""",
        "date": "2026-01-26",
        "recipients": []
    }
    
    fm = EmailFilterManager()
    mx_data = fm.extract_mx_data(sample_email)
    
    print("Test MX Extraction Results (English):")
    print(json.dumps(mx_data, indent=2, ensure_ascii=False))
    
    # Verify expected results
    expected = {
        "ticket_number": "INC0012345",
        "shop": "MX456",  # Should remove leading 0 and add MX
        "description": "Server is down and needs immediate attention"
    }
    
    print("\nExpected Results:")
    print(json.dumps(expected, indent=2))
    
    print("\nValidation:")
    print(f"Ticket number: {mx_data.get('ticket_number') == expected.get('ticket_number')}")
    print(f"Shop: {mx_data.get('shop') == expected.get('shop')}")
    print(f"Description extracted: {'description' in mx_data}")

if __name__ == "__main__":
    test_mx_extraction_english()