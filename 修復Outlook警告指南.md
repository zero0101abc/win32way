# 🔧 OUTLOOK 安全警告修復指南
# OUTLOOK SECURITY WARNING FIX GUIDE

## 📋 問題說明 (Problem Description)

**中文:**
Outlook 顯示安全警告："有程式存取電子郵件地址資訊或代表自己發送電子郵件"
這會阻止 Python 腳本讀取郵件。

**English:**
Outlook shows security warning: "A program is accessing your email addresses or sending email on your behalf"
This blocks the Python script from reading emails.

---

## ✅ 解決方案 (Solutions)

### 方案 1: 自動修復 (推薦) - AUTO FIX (Recommended)

**步驟 / Steps:**

1. **執行修復腳本 / Run fix script:**
```bash
python fix_outlook_warning.py
```

2. **選擇選項 1 / Choose option 1:**
```
Enter choice (1/2/3): 1
```

3. **關閉並重啟 Outlook / Close and restart Outlook**

4. **測試腳本 / Test script:**
```bash
python test_quick.py
```

✅ **完成！應該不再顯示警告 / Done! No more warnings!**

---

### 方案 2: 手動修復 (Registry) - MANUAL FIX

**步驟 / Steps:**

1. **雙擊檔案 / Double-click file:**
```
fix_outlook_security.reg
```

2. **點擊「是」確認 / Click "Yes" to confirm**

3. **關閉並重啟 Outlook / Close and restart Outlook**

4. **測試 / Test:**
```bash
python test_quick.py
```

**恢復原設定 / To undo:**
雙擊 / Double-click: `undo_outlook_security_fix.reg`

---

### 方案 3: 手動允許 (臨時) - MANUAL ALLOW (Temporary)

當警告出現時 / When warning appears:

1. ☑ **勾選「允許存取 10 分鐘」/ Check "Allow access for 10 minutes"**
2. ✅ **點擊「允許」/ Click "Allow"**
3. 🏃 **立即執行腳本 / Run script immediately:**
```bash
python test_quick.py
```

⏰ **注意：需在 10 分鐘內完成 / Must complete within 10 minutes**

---

## 🎯 快速開始 (Quick Start)

### 推薦流程 / Recommended Flow:

```bash
# 步驟 1: 修復 Outlook 安全警告
# Step 1: Fix Outlook security warning
python fix_outlook_warning.py
# 選擇 1 / Choose 1

# 步驟 2: 關閉 Outlook，然後重新開啟
# Step 2: Close Outlook, then reopen

# 步驟 3: 掃描郵件（應該無警告）
# Step 3: Scan emails (should have no warning)
python test_quick.py

# 步驟 4: 處理票證
# Step 4: Process tickets
python create_tickets.py

# 步驟 5: 開啟儀表板
# Step 5: Open dashboard
python server.py
# 瀏覽器開啟: http://localhost:8000/dashboard.html
```

---

## 📁 文件說明 (Files Explained)

| 檔案 / File | 用途 / Purpose |
|-------------|----------------|
| `fix_outlook_warning.py` | 🔧 自動偵測並修復警告 / Auto-detect and fix warning |
| `fix_outlook_security.reg` | 📝 手動修復（Registry 檔案）/ Manual fix (Registry file) |
| `undo_outlook_security_fix.reg` | ↩️ 恢復原設定 / Undo the fix |
| `FIX_OUTLOOK_SECURITY_WARNING.md` | 📖 完整說明文件 / Full documentation |

---

## ⚠️ 安全提醒 (Security Warning)

**中文:**
此修復會降低 Outlook 安全性，僅在信任的電腦上使用。
公司電腦可能被 IT 政策阻止。

**English:**
This fix reduces Outlook security. Only use on trusted computers.
Corporate computers may be blocked by IT policy.

**建議 / Recommendation:**
- ✅ 個人電腦 / Personal computer: 使用方案 1 或 2 / Use Option 1 or 2
- ⚠️ 公司電腦 / Company computer: 使用方案 3（手動允許）/ Use Option 3 (Manual allow)

---

## 🔍 故障排除 (Troubleshooting)

### 問題 1: 修復後仍出現警告 / Still shows warning after fix

**檢查 Office 版本 / Check Office version:**
```bash
python fix_outlook_warning.py
# 查看偵測到的版本 / See detected version
```

**可能原因 / Possible causes:**
- ❌ Outlook 未重啟 / Outlook not restarted
- ❌ 使用舊的 Office 版本（需修改 Registry 路徑）/ Old Office version (need different registry path)
- ❌ 公司 IT 政策阻止 / Blocked by company IT policy

**解決方法 / Solutions:**
1. 確認已關閉並重啟 Outlook / Confirm Outlook closed and restarted
2. 使用方案 3（手動允許）/ Use Option 3 (Manual allow)

---

### 問題 2: 無法執行 .reg 檔案 / Cannot run .reg file

**錯誤訊息 / Error message:**
```
Registry editing has been disabled by your administrator
```

**解決方法 / Solution:**
使用方案 3（手動允許）或聯絡 IT 部門
Use Option 3 (Manual allow) or contact IT department

---

### 問題 3: Python 腳本仍卡住 / Python script still stuck

**這是不同的問題 / This is a different issue**

參考文件 / See document: `FIX_EMAIL_SCAN_STUCK.md`

**快速修復 / Quick fix:**
```bash
python test_quick.py  # 使用快速掃描器 / Use fast scanner
```

---

## 📊 比較表 (Comparison Table)

| 方案 / Option | 難度 / Difficulty | 永久 / Permanent | 安全 / Security |
|--------------|------------------|-----------------|----------------|
| 方案 1 (自動) / Option 1 (Auto) | ⭐ 簡單 / Easy | ✅ 是 / Yes | ⚠️ 降低 / Reduced |
| 方案 2 (Registry) / Option 2 (Registry) | ⭐⭐ 中等 / Medium | ✅ 是 / Yes | ⚠️ 降低 / Reduced |
| 方案 3 (手動) / Option 3 (Manual) | ⭐⭐⭐ 簡單但重複 / Easy but repetitive | ❌ 否 / No | ✅ 保持 / Maintained |

---

## 💡 建議 (Recommendations)

### 個人電腦使用者 / Personal Computer Users:
✅ **使用方案 1（自動修復）/ Use Option 1 (Auto fix)**
- 一次設定，永久生效 / Set once, works forever
- 節省時間 / Saves time
- 適合日常使用 / Good for daily use

### 公司電腦使用者 / Company Computer Users:
✅ **使用方案 3（手動允許）/ Use Option 3 (Manual allow)**
- 保持安全設定 / Maintains security
- 不違反 IT 政策 / Doesn't violate IT policy
- 每次使用需允許 / Need to allow each time

### 開發測試 / Development/Testing:
✅ **使用方案 1，完成後恢復 / Use Option 1, undo when done**
```bash
# 修復 / Fix
python fix_outlook_warning.py  # Choose 1

# 開發工作... / Development work...

# 恢復 / Undo
python fix_outlook_warning.py  # Choose 2
```

---

## ✨ 完整工作流程 (Complete Workflow)

### 首次設定 / Initial Setup:

```bash
# 1. 修復 Outlook 警告
# 1. Fix Outlook warning
python fix_outlook_warning.py
選擇 1 / Choose 1

# 2. 重啟 Outlook
# 2. Restart Outlook

# 3. 測試掃描
# 3. Test scan
python test_quick.py
```

### 日常使用 / Daily Use:

```bash
# 掃描最新郵件（30 秒內完成）
# Scan recent emails (completes in 30s)
python test_quick.py

# 處理成票證
# Process into tickets
python create_tickets.py

# 開啟儀表板
# Open dashboard
python server.py
```

### 或使用儀表板 / Or Use Dashboard:

```bash
# 啟動伺服器
# Start server
python server.py

# 開啟瀏覽器
# Open browser
# http://localhost:8000/dashboard.html

# 點擊「Scan Emails」按鈕
# Click "Scan Emails" button
```

---

## 📞 需要協助？ (Need Help?)

### 查看所有文件 / Check all documents:

1. `FIX_OUTLOOK_SECURITY_WARNING.md` - 此文件的英文完整版 / Full English version of this doc
2. `FIX_EMAIL_SCAN_STUCK.md` - 解決掃描卡住問題 / Fix scanning stuck issue
3. `README.md` - 專案總覽 / Project overview

### 測試工具 / Testing Tools:

```bash
# 診斷 Outlook 連接
# Diagnose Outlook connection
python diagnose.py

# 檢查修復狀態
# Check fix status
python fix_outlook_warning.py
選擇 3 / Choose 3
```

---

## 🎉 成功指標 (Success Indicators)

修復成功後，您應該看到 / After successful fix, you should see:

✅ **腳本執行時無警告視窗 / No warning dialog when script runs**
✅ **test_quick.py 在 30 秒內完成 / test_quick.py completes in 30s**
✅ **outlook_emails.json 被創建 / outlook_emails.json is created**
✅ **儀表板可正常掃描 / Dashboard can scan normally**

---

**祝您使用順利！ / Good luck!** 🚀
