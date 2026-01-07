# Git Setup Guide - First Time

This guide will help you set up Git and push your code to a remote repository (GitHub/GitLab/etc.) for the first time.

## Prerequisites

1. **Install Git** (if not already installed):
   - Download from: https://git-scm.com/download/win
   - Or use: `winget install Git.Git` (Windows Package Manager)

2. **Create a GitHub account** (or GitLab/Bitbucket):
   - Go to: https://github.com
   - Sign up for a free account

## Step-by-Step Instructions

### Step 1: Open PowerShell/Terminal

Navigate to your project directory:
```powershell
cd C:\SalesIntel\repo\my_project\message-sender
```

### Step 2: Configure Git (First Time Only)

Set your name and email (used for commits):
```powershell
git config --global user.name "Your Name"
git config --global user.email "your.email@example.com"
```

### Step 3: Initialize Git Repository

Initialize a new Git repository in your project folder:
```powershell
git init
```

This creates a `.git` folder (hidden) that tracks your code.

### Step 4: Check Status

See what files will be added:
```powershell
git status
```

### Step 5: Add Files to Git

Add all files (respects .gitignore):
```powershell
git add .
```

Or add specific files:
```powershell
git add whatsapp_sender.py
git add README.md
git add requirements.txt
```

### Step 6: Make Your First Commit

Commit the files with a message:
```powershell
git commit -m "Initial commit: WhatsApp bulk message sender"
```

### Step 7: Create Remote Repository

**On GitHub:**
1. Go to https://github.com
2. Click the "+" icon → "New repository"
3. Name it (e.g., "whatsapp-message-sender")
4. **Don't** initialize with README (you already have files)
5. Click "Create repository"

### Step 8: Connect Local to Remote

Copy the repository URL from GitHub (HTTPS or SSH), then run:

**Using HTTPS:**
```powershell
git remote add origin https://github.com/yourusername/whatsapp-message-sender.git
```

**Using SSH (if you have SSH keys set up):**
```powershell
git remote add origin git@github.com:yourusername/whatsapp-message-sender.git
```

### Step 9: Push to Remote

Push your code to GitHub:
```powershell
git branch -M main
git push -u origin main
```

If prompted for credentials:
- **Username**: Your GitHub username
- **Password**: Use a Personal Access Token (not your GitHub password)
  - Create token: GitHub → Settings → Developer settings → Personal access tokens → Generate new token
  - Give it "repo" permissions

## Quick Reference Commands

### Daily Workflow

```powershell
# Check status
git status

# Add changes
git add .

# Commit changes
git commit -m "Description of changes"

# Push to remote
git push
```

### Useful Commands

```powershell
# See commit history
git log

# See what branch you're on
git branch

# Create a new branch
git checkout -b feature-name

# Switch branches
git checkout main

# Pull latest changes
git pull
```

## Troubleshooting

### If you get "fatal: not a git repository"
- Make sure you're in the project directory
- Run `git init` first

### If push is rejected
- Make sure you've committed your changes first
- Check if remote repository exists
- Try: `git push -u origin main --force` (⚠️ use carefully)

### If authentication fails
- Use Personal Access Token instead of password
- Or set up SSH keys for easier authentication

## Next Steps

1. **Create a README.md** (you already have one!)
2. **Add a LICENSE file** (if open source)
3. **Set up GitHub Actions** (for CI/CD, optional)
4. **Add collaborators** (if working in a team)

## Security Notes

⚠️ **Important**: Never commit:
- Passwords or API keys
- Personal data
- `.env` files with secrets
- `chrome_profile/` (contains browser data)

The `.gitignore` file is already set up to exclude these!
