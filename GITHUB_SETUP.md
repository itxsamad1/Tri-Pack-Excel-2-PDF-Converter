# GitHub Setup Instructions

## Step 1: Create Repository on GitHub

1. Go to https://github.com/new
2. Repository name: `Tri-Pack-Excel-2-PDF-Converter`
3. Description: "Excel to PDF converter for Tri-Pack Films pallet tags"
4. Choose Public or Private
5. **DO NOT** initialize with README, .gitignore, or license (we already have these)
6. Click "Create repository"

## Step 2: Push to GitHub

After creating the repository, run these commands:

```bash
git remote add origin https://github.com/YOUR_USERNAME/Tri-Pack-Excel-2-PDF-Converter.git
git branch -M main
git push -u origin main
```

Replace `YOUR_USERNAME` with your actual GitHub username.

## Alternative: If you prefer SSH

If you use SSH keys with GitHub:

```bash
git remote add origin git@github.com:YOUR_USERNAME/Tri-Pack-Excel-2-PDF-Converter.git
git branch -M main
git push -u origin main
```


