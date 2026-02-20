# OUTLOOK SECURITY WARNING FIX

## 🔴 PROBLEM
Outlook shows warning: "A program is accessing your email addresses or sending email on your behalf"
- Chinese: "我收到警告，指出有程式存取電子郵件地址資訊或代表自己發送電子郵件"
- This blocks the Python script from reading emails
- Script worked before, but now shows this warning

## ✅ SOLUTION OPTIONS

### Option 1: QUICK FIX - Trust the Python Script (Recommended)

When the warning appears:
1. ✅ **Check the box**: "Allow access for [X] minutes" (select 10 minutes)
2. ✅ **Click "Allow"**
3. ✅ **Run the script again within the time limit**

**Pros:** Easy, no system changes
**Cons:** Need to allow every time

---

### Option 2: Disable Outlook Security (PERMANENT FIX)

⚠️ **WARNING:** This reduces Outlook security. Only do this on trusted computers.

#### Method A: Registry Edit (Windows)

**Step 1**: Close Outlook completely

**Step 2**: Press `Win + R`, type `regedit`, press Enter

**Step 3**: Navigate to:
```
HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Security
```
Note: `16.0` is for Office 2016/2019/365. For Office 2013, use `15.0`

**Step 4**: Create new DWORD value:
- Name: `PromptOOMSend`
- Value: `2` (decimal)

**Step 5**: Create another DWORD value:
- Name: `PromptOOMAddressBookAccess`  
- Value: `2` (decimal)

**Step 6**: Create another DWORD value:
- Name: `PromptOOMAddressInformationAccess`
- Value: `2` (decimal)

**Step 7**: Close Registry Editor, restart Outlook

**To Undo:** Delete these registry values

---

### Option 3: AUTO-CLICK SOLUTION (Automated)

Use a modified script that automatically handles the warning:

```python
# test_auto_allow.py
import win32com.client
import win32gui
import win32con
import threading
import time

def auto_click_allow():
    """Auto-click the Allow button in Outlook security warning"""
    time.sleep(2)  # Wait for dialog to appear
    
    for _ in range(30):  # Try for 30 seconds
        # Find Outlook security dialog
        hwnd = win32gui.FindWindow(None, "Microsoft Outlook")
        if hwnd:
            # Find Allow button (usually the first button)
            try:
                # Press Enter to click default button
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
                print("Auto-clicked Allow button")
                break
            except:
                pass
        time.sleep(1)

# Start auto-clicker in background
threading.Thread(target=auto_click_allow, daemon=True).start()

# Rest of your scanning code here...
```

**Pros:** No manual clicking needed
**Cons:** May not work on all Outlook versions

---

### Option 4: USE REDEMPTION LIBRARY (Professional Solution)

Install the Redemption library to bypass Outlook security:

**Step 1**: Download Redemption
- Go to: http://www.dimastr.com/redemption/download.htm
- Download and install Redemption

**Step 2**: Install Python wrapper
```bash
pip install pywin32
```

**Step 3**: Use Redemption in your script:
```python
import win32com.client

# Create Redemption session instead of Outlook
session = win32com.client.Dispatch("Redemption.RDOSession")
session.Logon()

# Get inbox
inbox = session.GetDefaultFolder(6)
messages = inbox.Items

# No security warnings!
```

---

### Option 5: EXPORT TO PST (Alternative Approach)

Instead of accessing Outlook directly, export emails first:

**Step 1**: In Outlook, go to File → Open & Export → Import/Export

**Step 2**: Choose "Export to a file" → "Outlook Data File (.pst)"

**Step 3**: Select Inbox, export to `outlook_backup.pst`

**Step 4**: Use this script to read PST file:
```python
import win32com.client

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

# Open PST file (no security warning!)
pst_path = r"C:\path\to\outlook_backup.pst"
outlook.AddStore(pst_path)

# Access the PST folder
pst_folder = outlook.Folders.Item("outlook_backup")
# ... rest of code
```

---

### Option 6: CHANGE OUTLOOK TRUST CENTER SETTINGS

**Step 1**: Open Outlook

**Step 2**: Go to File → Options → Trust Center → Trust Center Settings

**Step 3**: Click "Programmatic Access"

**Step 4**: Select:
- ✅ "Never warn me about suspicious activity" (least secure)
- OR "Warn me about suspicious activity when my antivirus is inactive" (medium)

**Step 5**: Click OK, restart Outlook

⚠️ **Warning:** This reduces security for all programs, not just your script

---

### Option 7: ADMINISTRATOR GROUP POLICY (For IT Admins)

If you're on a corporate network, ask IT to add this policy:

**Group Policy Path:**
```
User Configuration → Administrative Templates → 
Microsoft Outlook 2016 → Security → Security Form Settings → 
Programmatic Security
```

Set to: "Automatically Approve" for trusted programs

---

## 🎯 RECOMMENDED SOLUTION FOR YOU

**Best approach (in order):**

### 1️⃣ **Quick Test** (Option 1):
Try clicking "Allow" for 10 minutes and run the script immediately to test

### 2️⃣ **Permanent Fix** (Option 2 - Registry):
If you trust your script and use your own computer, use the Registry method

### 3️⃣ **Professional** (Option 4 - Redemption):
If you need a production solution without security warnings

---

## 🛠️ IMPLEMENTATION - REGISTRY FIX SCRIPT

I'll create an automated script to apply the registry fix:

**File: `fix_outlook_security.reg`**
```reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Security]
"PromptOOMSend"=dword:00000002
"PromptOOMAddressBookAccess"=dword:00000002
"PromptOOMAddressInformationAccess"=dword:00000002
```

**To Apply:**
1. Save as `fix_outlook_security.reg`
2. Double-click the file
3. Click "Yes" to confirm
4. Restart Outlook

**To Undo:**
Delete the registry values or use this file:

```reg
Windows Registry Editor Version 5.00

[HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Security]
"PromptOOMSend"=-
"PromptOOMAddressBookAccess"=-
"PromptOOMAddressInformationAccess"=-
```

---

## 🔍 TROUBLESHOOTING

### Warning still appears after registry fix?

**Check your Office version:**
```bash
# Check Office version
reg query "HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" /v VersionToReport
```

Change registry path based on version:
- Office 2013: `15.0`
- Office 2016: `16.0`
- Office 2019/365: `16.0`

### Registry edit doesn't work?

Your company may have Group Policy blocking this. Try:
1. Option 1 (Manual allow)
2. Option 4 (Redemption)
3. Option 5 (Export to PST)

---

## ⚡ QUICK REFERENCE

| Solution | Security Impact | Ease | Permanent |
|----------|----------------|------|-----------|
| Manual Allow | ✅ Safe | ⭐⭐⭐⭐⭐ | ❌ No |
| Registry Edit | ⚠️ Reduces | ⭐⭐⭐⭐ | ✅ Yes |
| Auto-Click | ✅ Safe | ⭐⭐⭐ | ❌ No |
| Redemption | ✅ Safe | ⭐⭐ | ✅ Yes |
| PST Export | ✅ Safe | ⭐⭐⭐ | ❌ No |

---

## 📝 WHAT TO DO NOW

**Step 1**: Try the manual allow (Option 1)
```bash
# Run script, when warning appears:
# - Check "Allow access for 10 minutes"
# - Click Allow
python test_quick.py
```

**Step 2**: If that works, implement permanent fix (Option 2)
- Use the registry file I'll create
- Or manually edit registry

**Step 3**: Test the dashboard
```bash
python test_quick.py  # Should work without warning
python create_tickets.py
python server.py
```

Let me know which solution you'd like to implement!
