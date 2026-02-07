# SharePointFileRename
SharePoint PowerShell to rename files. This was purposely built to fix filenames after a Microsoft Point in Time restore from an Akira infection. 

[![PowerShell](https://img.shields.io/badge/PowerShell-7%2B-blue)](https://powershell.org)
[![PnP.PowerShell](https://img.shields.io/badge/PnP.PowerShell-latest-green)](https://pnp.github.io/powershell/)

This PowerShell script automates the renaming of SharePoint files that have a specific suffix (e.g., `.akira`) appended to their names. It was designed to complement Microsoft 365 point-in-time restores, where file contents are restored but file names may remain incorrect after events such as ransomware attacks.

> ⚠️ **Important:** This script **only renames files**. It does **not restore or repair damaged files**.

---

## Purpose

When performing a point-in-time restore of SharePoint sites, Microsoft 365 restores all file contents to the selected version. However, files that had their names changed during an incident (such as ransomware) **retain the altered names** after the restore.  

This script provides a safe and automated way to **remove unwanted suffixes** and restore the original file names **without affecting the file contents**.

---

## Features

- Uses **app-only authentication** for secure, non-interactive access to SharePoint Online.
- Supports **single libraries, specific folders, or multiple sites**.
- Filters files based on a **specific suffix** (e.g., `.akira`).
- Supports **dry-run mode** (`-WhatIf`) for safe testing.
- Generates a **log file** for auditing renamed files.
- Compatible with **Microsoft Teams sites** and known SharePoint site URLs.

---

## Prerequisites

- **PowerShell 7+** (recommended)
- **PnP.PowerShell** module installed:

```powershell
Install-Module PnP.PowerShell -Scope CurrentUser
