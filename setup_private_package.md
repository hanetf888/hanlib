# Setting Up a Private Shared Python Package

This guide walks you through creating a private pip-installable Python package to share modules like `database.py` and `utils.py` across multiple projects.

**Note:** All commands in this guide are for PowerShell on Windows.

---

## Step 1: Create the Package Directory Structure

Create a new folder for your shared package with this structure:

```
hanetf-common/
├── pyproject.toml
├── README.md
├── src/
│   └── hanlib/
│       ├── __init__.py
│       ├── database.py
│       ├── utils.py
│       └── hanconfig.py
```

**Commands:**
```powershell
mkdir hanetf-common
cd hanetf-common
New-Item -ItemType Directory -Force -Path src\hanlib
```

---

## Step 2: Create pyproject.toml

Create `hanetf-common/pyproject.toml`:

```toml
[build-system]
requires = ["setuptools>=61.0"]
build-backend = "setuptools.build_meta"

[project]
name = "hanetf-common"
version = "1.0.0"
description = "Shared utilities for HANetf projects"
readme = "README.md"
requires-python = ">=3.9"
dependencies = [
    "pandas",
    "psycopg2-binary",
    "sqlalchemy",
    "python-dotenv",
    "numpy",
    "pyodbc",
]

[project.optional-dependencies]
dev = [
    "pytest",
]

[tool.setuptools.packages.find]
where = ["src"]
```

---

## Step 3: Create __init__.py

Create `hanetf-common/src/hanlib/__init__.py`:

```python
from . import database
from . import utils

__version__ = "1.0.0"
```

---

## Step 4: Copy Your Shared Modules

Copy your existing modules into the package:

```powershell
# From your geoffrey project directory
Copy-Item lib\database.py ..\hanetf-common\src\hanlib\
Copy-Item lib\utils.py ..\hanetf-common\src\hanlib\
Copy-Item lib\hanconfig.py ..\hanetf-common\src\hanlib\
```

---

## Step 5: Create a Private GitHub Repository

1. Go to https://github.com/new
2. Repository name: `hanetf-common`
3. **Select "Private"**
4. Click "Create repository"

---

## Step 6: Push to GitHub

```powershell
cd hanetf-common
git init
git add .
git commit -m "Initial commit: shared database and utils modules"
git branch -M main
git remote add origin git@github.com:hanetf888/hanetf-common.git
git push -u origin main
```

---

## Step 7: Set Up SSH Key (if not already done)

### SSH Key Location

On Windows, SSH keys are stored in your user profile directory:

| Key Type | Location |
|----------|----------|
| Private key | `C:\Users\<YourUsername>\.ssh\id_rsa` |
| Public key | `C:\Users\<YourUsername>\.ssh\id_rsa.pub` |

In PowerShell, you can reference these paths using `$env:USERPROFILE`:
- Private key: `$env:USERPROFILE\.ssh\id_rsa`
- Public key: `$env:USERPROFILE\.ssh\id_rsa.pub`

### Check if you have an SSH key

```powershell
Test-Path "$env:USERPROFILE\.ssh\id_rsa.pub"
```

If this returns `True`, you already have a key. To view it:
```powershell
Get-Content "$env:USERPROFILE\.ssh\id_rsa.pub"
```

### Generate a new SSH key (if needed)

```powershell
ssh-keygen -t rsa -b 4096 -C "your_email@example.com"
```

When prompted for a location, press Enter to accept the default (`C:\Users\<YourUsername>\.ssh\id_rsa`).

### Add the public key to GitHub

1. Copy the key to clipboard:
   ```powershell
   Get-Content "$env:USERPROFILE\.ssh\id_rsa.pub" | Set-Clipboard
   ```
2. Go to GitHub → Settings → SSH and GPG keys → New SSH key
3. Paste and save

---

## Step 8: Install the Package in Your Projects

### Option A: Install directly from GitHub (SSH)
```powershell
pip install git+ssh://git@github.com/hanetf888/hanlib.git
```

### Option B: Install a specific version/tag
```powershell
# First, create a tag in your hanetf-common repo
git tag v1.0.0
git push origin v1.0.0

# Then install that version
pip install git+ssh://git@github.com/hanetf888/hanlib.git@v1.0.0
```

### Option C: Add to requirements.txt
```
git+ssh://git@github.com/hanetf888/hanlib.git@v1.0.0
```

---

## Step 9: Update Your Project Code

Change your imports from:
```python
import database as han_db
import utils
import hanconfig
```

To:
```python
from hanlib import database as han_db
from hanlib import utils
from hanlib import hanconfig
```

---

## Step 10: Update and Versioning

When you make changes to the shared package:

1. Update the code in `hanetf-common`
2. Update version in `pyproject.toml` and `__init__.py`
3. Commit and push:
   ```powershell
   git add .
   git commit -m "Description of changes"
   git tag v1.1.0
   git push origin main --tags
   ```

4. Update in your projects:
   ```powershell
   pip install --upgrade git+ssh://git@github.com/hanetf888/hanetf-common.git@v1.1.0
   ```

---

## Troubleshooting

### SSH Permission Denied

Test SSH connection:
```powershell
ssh -T git@github.com
```

If it fails, ensure the SSH agent service is running:
```powershell
# Check if ssh-agent service exists and its status
Get-Service ssh-agent

# Start the service (requires admin PowerShell)
Start-Service ssh-agent

# Set it to start automatically (requires admin PowerShell)
Set-Service ssh-agent -StartupType Automatic

# Add your key to the agent
ssh-add "$env:USERPROFILE\.ssh\id_rsa"
```

**Note:** If the ssh-agent service doesn't exist, you may need to install OpenSSH:
1. Go to Settings → Apps → Optional Features
2. Add "OpenSSH Client" (and optionally "OpenSSH Server")

### Package Not Found After Install

Ensure the package structure is correct:
```powershell
pip show hanetf-common
```

### Editable Install for Development

If you're actively developing the shared package:
```powershell
cd hanetf-common
pip install -e .
```

---

## Summary

| Task | Command |
|------|---------|
| Install package | `pip install git+ssh://git@github.com/hanetf888/hanetf-common.git` |
| Install specific version | `pip install git+ssh://git@github.com/hanetf888/hanetf-common.git@v1.0.0` |
| Upgrade package | `pip install --upgrade git+ssh://git@github.com/hanetf888/hanetf-common.git` |
| Uninstall | `pip uninstall hanetf-common` |
| Check installed version | `pip show hanetf-common` |
| Check SSH key exists | `Test-Path "$env:USERPROFILE\.ssh\id_rsa.pub"` |
| View SSH public key | `Get-Content "$env:USERPROFILE\.ssh\id_rsa.pub"` |
| Copy SSH key to clipboard | `Get-Content "$env:USERPROFILE\.ssh\id_rsa.pub" \| Set-Clipboard` |
