# Easy PDF Watcher

A custom Windows workflow tool for automatic PDF detection, work order extraction, file renaming, and export routing.

## Overview

Easy PDF Watcher was built to simplify a repetitive PDF workflow. Instead of manually renaming exported files every time, this tool watches a selected folder, detects new PDF files, extracts the work order from the original file name, and opens a review window before saving the final renamed file to an export folder.

The tool is designed for practical day-to-day use and includes a simple setup flow so each user can choose their own folders without editing code.

## Key Features

- Watches a selected folder for new PDF files
- Detects work order codes from the original file name
- Generates a standardized output name using this format:

`YYMMDD_HHMM_WO.pdf`

Example:

`260414_1544_R3805.pdf`

- Opens a review window before applying the rename
- Lets the user edit the detected work order if needed
- Sends the renamed PDF to a chosen export folder
- Includes a setup utility for folder configuration
- Supports automatic startup through the VBS launcher
- Includes a branded UI and notification flow
![Easy PDF Watcher visual selection](main_files/Easy%20PDF%20Watcher%20-%20visual%20selection.png)

## How It Works

1. The user runs the setup tool and chooses:
   - a **watch folder**
   - an **export folder**

2. The user launches the main watcher.

3. When a new PDF appears in the watch folder:
   - the tool detects it
   - reads the work order from the original file name
   - generates a suggested final name
   - shows a review window for confirmation

4. After confirmation:
   - the file is renamed
   - moved to the export folder
   - and a notification appears indicating the file is ready

## File Naming Logic

The tool builds file names in this format:

`YYMMDD_HHMM_WO.pdf`

The work order is extracted from the original PDF name when possible.

Examples:

- `r3805 lves_romito22.pdf` → `260414_1544_R3805.pdf`
- `n1173 project export.pdf` → `260414_1544_N1173.pdf`

If no valid work order is found, the tool uses:

`EDIT_ME`

This allows the user to manually fix the value before applying the final name.

## Project Structure

```text
easy-pdf-watcher/
├─ main_files/
│  ├─ main_pdf_easy_watcher.ps1
│  ├─ setup_pdf_easy_watcher.ps1
│  └─ marca_sola.png
├─ vbs_files/
│  ├─ main_pdf_easy_watcher.vbs
│  └─ setup_pdf_easy_watcher.vbs
├─ LICENSE
└─ README.md
