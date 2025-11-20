# Implementation Plan

- [x] 1. Implement campus name normalization function
  - Create `normalizeCampusName()` function that takes a raw campus name string
  - Strip teacher name suffix by removing everything after and including " - "
  - Replace space-separated numbers (" 1", " 2", " 3", etc.) with "# " format (" #1", " #2", " #3")
  - Convert result to lowercase for case-insensitive comparison
  - Return normalized campus name string
  - _Requirements: 2.1, 2.2, 2.3, 3.1, 3.2, 3.3_

- [x] 2. Implement source sheet data reader function
  - Create `readSourceSheetData()` function that accepts a sheet name parameter
  - Get sheet reference using `getSheetByName()`
  - Return empty array if sheet doesn't exist (log warning)
  - Get last row with data using `getLastRow()`
  - Return empty array if no data rows exist (lastRow < 2)
  - Read campus names from column B (rows 2 to lastRow) in single batch operation
  - Read count values from column D (rows 2 to lastRow) in single batch operation
  - Build array of objects with campus and count properties
  - Convert empty or invalid count values to 0
  - Return array of campus-count pairs
  - _Requirements: 1.1, 1.2, 1.3, 5.1, 5.2, 5.4_

- [x] 3. Implement main campus count aggregation function
  - Create `aggregateCampusCounts()` function
  - Get active spreadsheet reference
  - Get ALE Counts sheet reference
  - Display error and return if ALE Counts sheet doesn't exist
  - Read campus names from column D rows 4-212 in single batch operation
  - Build campus lookup map with normalized names as keys and row numbers as values
  - Call `readSourceSheetData()` for ES sheet and process results
  - Call `readSourceSheetData()` for MS sheet and process results
  - Call `readSourceSheetData()` for HS sheet and process results
  - For each source data record, normalize campus name and accumulate count in lookup map
  - Build output array with aggregated counts aligned to ALE Counts rows
  - Write all counts to column E (rows 4-212) in single batch operation
  - Count number of campuses with non-zero values
  - Display success message with count of updated campuses
  - Return number of campuses updated
  - _Requirements: 1.4, 2.4, 2.5, 3.4, 4.4, 5.3, 5.5_

- [x] 4. Integrate aggregation into existing extraction functions
  - Add call to `aggregateCampusCounts()` at the end of `extractESGradeLevelValue()`
  - Add call to `aggregateCampusCounts()` at the end of `extractMSGradeLevelValue()`
  - Add call to `aggregateCampusCounts()` at the end of `extractHSGradeLevelValue()`
  - Ensure aggregation runs after grade level extraction completes
  - _Requirements: 4.1, 4.2, 4.3, 4.5_
