#!/usr/bin/env python3
"""
Diagnostic script to find where test.py hangs
"""
import win32com.client
import time

print("Step 1: Testing Outlook connection...")
start = time.time()
try:
    outlook = win32com.client.Dispatch("Outlook.Application")
    print(f"  OK - Outlook app created ({time.time()-start:.1f}s)")
except Exception as e:
    print(f"  FAILED: {e}")
    exit(1)

print("\nStep 2: Getting MAPI namespace...")
start = time.time()
try:
    namespace = outlook.GetNamespace("MAPI")
    print(f"  OK - Namespace created ({time.time()-start:.1f}s)")
except Exception as e:
    print(f"  FAILED: {e}")
    exit(1)

print("\nStep 3: Getting inbox folder...")
start = time.time()
try:
    inbox = namespace.GetDefaultFolder(6)
    print(f"  OK - Inbox retrieved ({time.time()-start:.1f}s)")
except Exception as e:
    print(f"  FAILED: {e}")
    exit(1)

print("\nStep 4: Getting messages collection...")
start = time.time()
try:
    messages = inbox.Items
    print(f"  OK - Messages collection created ({time.time()-start:.1f}s)")
except Exception as e:
    print(f"  FAILED: {e}")
    exit(1)

print("\nStep 5: Counting messages...")
start = time.time()
try:
    count = messages.Count
    print(f"  OK - Found {count} messages ({time.time()-start:.1f}s)")
except Exception as e:
    print(f"  FAILED: {e}")
    exit(1)

print(f"\n*** PROBLEM IS IN STEP 6 OR LATER ***")

print("\nStep 6: Sorting messages by date...")
start = time.time()
try:
    messages.Sort("[ReceivedTime]", True)
    print(f"  OK - Sorted ({time.time()-start:.1f}s)")
except Exception as e:
    print(f"  FAILED: {e}")
    exit(1)

print("\nStep 7: Accessing first message...")
start = time.time()
try:
    msg = messages[0]
    print(f"  OK - Got message object ({time.time()-start:.1f}s)")
except Exception as e:
    print(f"  FAILED: {e}")
    exit(1)

print("\nStep 8: Getting subject...")
start = time.time()
try:
    subject = msg.Subject
    print(f"  OK - Subject: {subject[:50]}... ({time.time()-start:.1f}s)")
except Exception as e:
    print(f"  FAILED: {e}")
    exit(1)

print("\nStep 9: Getting body (THIS OFTEN HANGS)...")
start = time.time()
try:
    body = msg.Body
    print(f"  OK - Body length: {len(body)} ({time.time()-start:.1f}s)")
except Exception as e:
    print(f"  FAILED: {e}")
    exit(1)

print("\nStep 10: Getting recipients...")
start = time.time()
try:
    recipients = []
    for r in msg.Recipients:
        recipients.append(r.Name)
    print(f"  OK - Found {len(recipients)} recipients ({time.time()-start:.1f}s)")
except Exception as e:
    print(f"  FAILED: {e}")
    exit(1)

print("\n" + "="*60)
print("ALL STEPS PASSED!")
print("="*60)
print(f"\nThe problem is NOT in Outlook access.")
print(f"It must be in the loop iteration or filter processing.")
