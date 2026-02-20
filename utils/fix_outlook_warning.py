#!/usr/bin/env python3
"""
Automatically fix Outlook security warning
This script modifies the Windows Registry to allow Python to access Outlook
"""

import winreg
import sys

def check_office_version():
    """Detect installed Office version"""
    versions = {
        "16.0": "Office 2016/2019/365",
        "15.0": "Office 2013",
        "14.0": "Office 2010"
    }
    
    for version, name in versions.items():
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                f"Software\\Microsoft\\Office\\{version}\\Outlook\\Security",
                0,
                winreg.KEY_READ
            )
            winreg.CloseKey(key)
            return version, name
        except:
            continue
    
    return None, None

def check_security_settings(version):
    """Check if security settings are already disabled"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            f"Software\\Microsoft\\Office\\{version}\\Outlook\\Security",
            0,
            winreg.KEY_READ
        )
        
        settings = {}
        for name in ["PromptOOMSend", "PromptOOMAddressBookAccess", "PromptOOMAddressInformationAccess"]:
            try:
                value, _ = winreg.QueryValueEx(key, name)
                settings[name] = value
            except:
                settings[name] = None
        
        winreg.CloseKey(key)
        return settings
    except:
        return {}

def apply_fix(version):
    """Apply the registry fix"""
    try:
        # Open or create the Security key
        key = winreg.CreateKey(
            winreg.HKEY_CURRENT_USER,
            f"Software\\Microsoft\\Office\\{version}\\Outlook\\Security"
        )
        
        # Set the values to disable prompts
        winreg.SetValueEx(key, "PromptOOMSend", 0, winreg.REG_DWORD, 2)
        winreg.SetValueEx(key, "PromptOOMAddressBookAccess", 0, winreg.REG_DWORD, 2)
        winreg.SetValueEx(key, "PromptOOMAddressInformationAccess", 0, winreg.REG_DWORD, 2)
        
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Error: {e}")
        return False

def remove_fix(version):
    """Remove the registry fix (restore security)"""
    try:
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            f"Software\\Microsoft\\Office\\{version}\\Outlook\\Security",
            0,
            winreg.KEY_SET_VALUE
        )
        
        for name in ["PromptOOMSend", "PromptOOMAddressBookAccess", "PromptOOMAddressInformationAccess"]:
            try:
                winreg.DeleteValue(key, name)
            except:
                pass
        
        winreg.CloseKey(key)
        return True
    except Exception as e:
        print(f"Error: {e}")
        return False

def main():
    print("="*60)
    print("OUTLOOK SECURITY WARNING FIX")
    print("="*60)
    
    # Detect Office version
    print("\n[1/4] Detecting Office version...")
    version, name = check_office_version()
    
    if not version:
        print("ERROR: Could not detect Office/Outlook installation")
        print("Make sure Outlook is installed")
        return
    
    print(f"Found: {name} (version {version})")
    
    # Check current settings
    print("\n[2/4] Checking current security settings...")
    settings = check_security_settings(version)
    
    if settings:
        print("Current settings:")
        for name, value in settings.items():
            status = "Disabled (2)" if value == 2 else "Enabled (default)" if value is None else f"Custom ({value})"
            print(f"  {name}: {status}")
    
    # Ask user what to do
    print("\n[3/4] Choose action:")
    print("  1. Apply fix (disable security warnings)")
    print("  2. Remove fix (restore security warnings)")
    print("  3. Check only (no changes)")
    
    choice = input("\nEnter choice (1/2/3): ").strip()
    
    print("\n[4/4] Processing...")
    
    if choice == "1":
        print("Applying fix...")
        if apply_fix(version):
            print("✓ SUCCESS! Security warnings disabled")
            print("\nNext steps:")
            print("  1. Close Outlook if it's running")
            print("  2. Restart Outlook")
            print("  3. Run: python test_quick.py")
            print("\nThe script should now work without warnings!")
        else:
            print("✗ FAILED to apply fix")
            print("Try running this script as Administrator")
    
    elif choice == "2":
        print("Removing fix...")
        if remove_fix(version):
            print("✓ SUCCESS! Security warnings restored")
            print("\nOutlook security is back to default")
        else:
            print("✗ FAILED to remove fix")
    
    elif choice == "3":
        print("No changes made")
    
    else:
        print("Invalid choice")
    
    print("\n" + "="*60)

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nERROR: {e}")
        print("\nTry running as Administrator:")
        print("  Right-click Command Prompt → Run as Administrator")
        print("  Then run: python fix_outlook_warning.py")
