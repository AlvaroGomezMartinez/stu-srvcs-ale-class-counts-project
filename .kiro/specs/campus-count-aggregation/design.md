# Design Document

## Overview

This design implements an automated campus count aggregation system that consolidates enrollment data from three level-specific sheets (ES, MS, HS) into a unified ALE Counts reporting sheet. The solution addresses the core challenge of matching campus names across sheets despite formatting inconsistencies, using a normalization and fuzzy matching approach.

The system integrates seamlessly with the existing Google Apps Script infrastructure by hooking into the existing grade level extraction functions, ensuring counts are automatically updated whenever source data is refreshed.

## Architecture

### High-Level Flow

```
User triggers grade level extraction (ES/MS/HS)
    ↓
Extract grade level values (existing functionality)
    ↓
Trigger campus count aggregation (new functionality)
    ↓
Read campus names from ALE Counts sheet (rows 4-212)
    ↓
Read count data from ES, MS, HS sheets
    ↓
Normalize and match campus names
    ↓
Aggregate matching counts
    ↓
Write totals to ALE Counts column E
    ↓
Display confirmation message
```

### Integration Points

The aggregation function will be called at the end of:
- `extractESGradeLevelValue()`
- `extractMSGradeLevelValue()`
- `extractHSGradeLevelValue()`

This ensures the ALE Counts sheet is updated immediately after source data is refreshed.

## Components and Interfaces

### 1. Campus Name Normalizer

**Purpose:** Convert campus names from Source Sheets into a standardized format that matches ALE Counts naming conventions

**Function Signature:**
```javascript
/**
 * Normalize a campus name from source sheets to match ALE Counts format
 * @param {string} sourceName - Raw campus name from ES/MS/HS sheet
 * @returns {string} Normalized campus name
 */
function normalizeCampusName(sourceName)
```

**Normalization Rules:**
1. Remove teacher name suffix (everything after and including " - ")
2. Trim whitespace
3. Replace space-separated numbers with "#" format:
   - " 1" → " #1"
   - " 2" → " #2"
   - " 3" → " #3"
   - etc.
4. Preserve "AU" designations
5. Preserve parenthetical level indicators like "(MS)", "(HS)"
6. Convert to lowercase for comparison

**Examples:**
- "Bernal 1 - Tracey Sorrell" → "bernal #1"
- "Brandeis AU - Stacey Wallace" → "brandeis au"
- "Holmgreen - Michael Cline" → "holmgreen"
- "Clark ( 3 Periods ) - Karen Pumphrey" → "clark ( 3 periods )"

### 2. Campus Count Aggregator

**Purpose:** Main orchestration function that reads data, matches campuses, and writes results

**Function Signature:**
```javascript
/**
 * Aggregate campus counts from ES, MS, HS sheets into ALE Counts sheet
 * @returns {number} Number of campuses updated
 */
function aggregateCampusCounts()
```

**Process:**
1. Get reference to active spreadsheet
2. Get reference to ALE Counts sheet
3. Validate ALE Counts sheet exists
4. Read campus names from column D (rows 4-212)
5. Build campus lookup map (normalized name → row number)
6. Process each source sheet (ES, MS, HS):
   - Read campus names from column B (starting row 2)
   - Read count values from column D (starting row 2)
   - Normalize each campus name
   - Accumulate counts in lookup map
7. Write aggregated counts to column E of ALE Counts
8. Return count of updated campuses

### 3. Source Sheet Data Reader

**Purpose:** Extract campus names and counts from a single source sheet

**Function Signature:**
```javascript
/**
 * Read campus count data from a source sheet
 * @param {string} sheetName - Name of the sheet (ES, MS, or HS)
 * @returns {Array<{campus: string, count: number}>} Array of campus-count pairs
 */
function readSourceSheetData(sheetName)
```

**Process:**
1. Get sheet by name
2. If sheet doesn't exist, log warning and return empty array
3. Get last row with data
4. If no data rows (lastRow < 2), return empty array
5. Read column B (campus names) and column D (counts) from row 2 to lastRow
6. For each row:
   - Skip if campus name is empty
   - Parse count value (default to 0 if empty/invalid)
   - Add to results array
7. Return results array

### 4. Integration Hook

**Purpose:** Modify existing extraction functions to trigger aggregation

**Implementation:**
Add the following line at the end of each extraction function:
```javascript
aggregateCampusCounts();
```

Modified functions:
- `extractESGradeLevelValue()`
- `extractMSGradeLevelValue()`
- `extractHSGradeLevelValue()`

## Data Models

### Campus Lookup Map
```javascript
{
  "normalized-campus-name": {
    row: number,        // Row number in ALE Counts sheet
    originalName: string, // Original name from ALE Counts
    totalCount: number   // Accumulated count from source sheets
  }
}
```

### Source Data Record
```javascript
{
  campus: string,  // Raw campus name from source sheet
  count: number    // Count value (0 if empty/invalid)
}
```

## Error Handling

### Missing ALE Counts Sheet
- **Detection:** Check if `getSheetByName("ALE Counts")` returns null
- **Action:** Display error alert and return early
- **Message:** "Error: ALE Counts sheet not found"

### Missing Source Sheet
- **Detection:** Check if `getSheetByName(sheetName)` returns null in `readSourceSheetData()`
- **Action:** Log warning, return empty array, continue processing other sheets
- **Message:** `Warning: ${sheetName} sheet not found, skipping`

### Invalid Count Values
- **Detection:** Check if count value is not a number or is empty
- **Action:** Default to 0, log warning
- **Message:** `Warning: Invalid count value in ${sheetName} row ${rowNum}, defaulting to 0`

### No Matching Campuses
- **Detection:** After processing all source sheets, check if any campuses have totalCount > 0
- **Action:** Display warning message
- **Message:** "Warning: No matching campuses found between source sheets and ALE Counts"

### Multiple Matches (Same Campus in Multiple Sheets)
- **Detection:** When normalized campus name already exists in lookup map with count > 0
- **Action:** Sum the counts (this is expected behavior for campuses that span multiple levels)
- **Message:** No warning needed (this is normal)

## Testing Strategy

Manual verification will be performed by running the extraction functions and confirming that counts are correctly aggregated in the ALE Counts sheet.

## Performance Considerations

### Batch Operations
- Read all campus names from ALE Counts in a single `getRange()` call
- Read all source data in a single `getRange()` call per sheet
- Write all aggregated counts in a single `setValues()` call

### Expected Performance
- Campus lookup map: ~160 entries (O(1) lookup)
- Source sheet reads: 3 sheets × ~50-100 rows each
- Total execution time: < 5 seconds

### Optimization Notes
- Avoid individual cell reads/writes in loops
- Use array operations for data processing
- Minimize spreadsheet API calls

## Special Cases

### Multi-Level Campuses
Campuses like "Holmgreen" appear in multiple sheets with level indicators:
- "Holmgreen (MS)" in MS sheet
- "Holmgreen (HS) #1", "Holmgreen (HS) #2", "Holmgreen (HS) #3" in HS sheet
- "Reddix (multi-level)" in special programs

**Handling:** Preserve parenthetical indicators during normalization to ensure correct matching

### Multi-Period Campuses
Some campuses have period indicators:
- "Clark ( 3 Periods )" in HS sheet

**Handling:** Preserve parenthetical content during normalization

### AU (Alternative Unit) Campuses
Many campuses have AU designation:
- "Boone AU", "Fernandez AU", "Lewis AU", etc.

**Handling:** Preserve "AU" during normalization, ensure it's included in matching logic

### Vacant Positions
Some source sheet entries show "Vacant" as teacher name:
- "Colonies North - Vacant"
- "Michael 2 - Vacant"

**Handling:** Teacher names are stripped regardless of content, so "Vacant" is removed like any other teacher name

### Template Rows
Source sheets contain "Template" rows with 0 counts:
- These should be processed normally but will likely not match any ALE Counts campus
- No special handling needed

## Dependencies

- Google Apps Script SpreadsheetApp API
- Existing CONFIGS object (no changes needed)
- Existing menu structure (no changes needed)
- Existing extraction functions (modified to call aggregation)
