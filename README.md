# ALE Class Counts Automation

Google Apps Script project that automates the collection and aggregation of Alternative Learning Experience (ALE) class counts across Elementary, Middle, and High School campuses.

## Overview

This script streamlines the process of gathering student enrollment data from multiple Google Sheets spread across different school campuses and consolidating them into a single "ALE Counts" sheet. It eliminates manual data entry by automatically extracting teacher reported class count values from individual campus spreadsheets.

## Features

- **Automated Data Collection**: Scans Drive folders for campus spreadsheets and extracts enrollment counts
- **Multi-Level Support**: Handles Elementary (ES), Middle (MS), and High School (HS) data separately or together
- **Campus Mapping**: Uses predefined mappings to match spreadsheet IDs with campus names
- **Batch Processing**: Efficiently processes multiple spreadsheets using batch operations
- **Error Handling**: Logs warnings for missing data, permission issues, and invalid values
- **Custom Menu**: Provides an intuitive UI menu in Google Sheets for easy operation

## Project Structure

```
├── Code.js              # Main script with data extraction and aggregation logic
├── CampusMapping.js     # Campus name to spreadsheet ID mappings for ES/MS/HS
├── appsscript.json      # Apps Script manifest with OAuth scopes
├── .clasp.json          # Clasp configuration for local development
└── README.md            # This file
```

## How It Works

### Workflow

1. **Get Spreadsheet IDs**: Scans a configured Drive folder and lists all Google Sheets with their IDs
2. **Get Counts**: Opens each spreadsheet, finds the "Current Grade Level" header, and extracts the enrollment value
3. **Aggregate Data**: Matches spreadsheet IDs to campus names and writes totals to the "ALE Counts" sheet

### Menu Structure

The script adds an "Update Counts" menu to your Google Sheet with three sub-menus:

- **Elementary School**
  - 1. Get Spreadsheet IDs
  - 2. Get Counts
- **Middle School**
  - 1. Get Spreadsheet IDs
  - 2. Get Counts
- **High School**
  - 1. Get Spreadsheet IDs
  - 2. Get Counts

## Configuration

The script uses folder IDs defined in the `CONFIGS` object in `Code.js`:

```javascript
const CONFIGS = {
  ES: { folderId: "...", sheetName: "ES" },
  MS: { folderId: "...", sheetName: "MS" },
  HS: { folderId: "...", sheetName: "HS" }
};
```

Campus mappings are maintained in `CampusMapping.js` with three Map objects:
- `elementarySchoolCampusMap`
- `middleSchoolCampusMap`
- `highSchoolCampusMap`

## Requirements

- Google Workspace account with access to:
  - Google Sheets API
  - Google Drive API
- Appropriate permissions to access campus spreadsheets
- Drive folders containing campus spreadsheets must be accessible

## License

MIT