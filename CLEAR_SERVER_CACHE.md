# Clear Server Cache

This guide explains how to clear the server cache for the ticket management system.

## Quick Commands

### Clear Python Cache

Run from the project root directory (`win32way-`):

```powershell
Get-ChildItem -Path . -Recurse -Directory -Filter __pycache__ | Remove-Item -Recurse -Force
```

### Clear All Cache Files

Remove `.pyc` files, `.pytest_cache`, and `.eggs` directories:

```powershell
Get-ChildItem -Path . -Recurse -Filter "*.pyc" -File | Remove-Item -Force
Get-ChildItem -Path . -Recurse -Filter ".pytest_cache" -Directory | Remove-Item -Recurse -Force
Get-ChildItem -Path . -Recurse -Filter ".eggs" -Directory | Remove-Item -Recurse -Force
```

## What Gets Cleared

| Item | Description |
|------|-------------|
| `__pycache__/` | Python bytecode cache directories |
| `*.pyc` | Compiled Python files |
| `.pytest_cache/` | Pytest cache directory |
| `.eggs/` | Eggs directory |

## When to Clear Cache

- **Before starting the server** - Ensures fresh Python imports
- **After updating dependencies** - Clears old module versions
- **On server errors** - May resolve caching issues
- **Before testing** - Ensures clean test environment

## Server Management

### Start Server
```powershell
python -m backend.server
```

### Stop Server
Kill all Python processes:
```powershell
Get-Process | Where-Object {$_.ProcessName -like "*python*"} | Stop-Process -Force
```

## Notes

- Cache files are automatically regenerated when needed
- Clearing cache is safe and will not lose any data
- Cache clearing helps with development and debugging


ctrl + alt + i /L