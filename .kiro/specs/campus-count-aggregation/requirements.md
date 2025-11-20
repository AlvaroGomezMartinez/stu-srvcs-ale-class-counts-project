# Requirements Document

## Introduction

This feature enables automated aggregation of campus enrollment counts from three separate tracking sheets (ES, MS, and HS) into a consolidated reporting sheet (ALE Counts). The system must handle campus name variations between source and destination sheets by implementing intelligent name matching logic that accounts for formatting differences such as "#" vs numeric indicators and teacher name suffixes.

## Glossary

- **ALE Counts Sheet**: The destination reporting sheet containing consolidated campus data with headers in row 3: "Teacher, Tier, Cluster, Campus, Total Enrolled, ALE IAs, Notes"
- **Source Sheets**: The three level-specific sheets (ES, MS, HS) containing campus count data with headers in row 1: "Spreadsheet ID, Campus, Error Log, Count"
- **Campus Name Mapping**: The process of matching campus names between Source Sheets and ALE Counts Sheet despite formatting differences
- **Count Column**: Column D in Source Sheets (ES, MS, HS) containing numeric enrollment values
- **Total Enrolled Column**: Column E in ALE Counts Sheet where aggregated counts are written
- **Campus Identifier**: The base campus name without teacher names or formatting variations (e.g., "Bernal 1" matches "Bernal #1")
- **Google Apps Script System**: The execution environment for the spreadsheet automation

## Requirements

### Requirement 1

**User Story:** As a data administrator, I want to aggregate campus counts from ES, MS, and HS sheets into the ALE Counts sheet, so that I can view consolidated enrollment data in one location

#### Acceptance Criteria

1. WHEN the aggregation function is triggered, THE Google Apps Script System SHALL read all values from column D (Count) of the ES sheet
2. WHEN the aggregation function is triggered, THE Google Apps Script System SHALL read all values from column D (Count) of the MS sheet
3. WHEN the aggregation function is triggered, THE Google Apps Script System SHALL read all values from column D (Count) of the HS sheet
4. WHEN count values are retrieved from Source Sheets, THE Google Apps Script System SHALL write matching values to column E (Total Enrolled) of the ALE Counts Sheet starting at row 4

### Requirement 2

**User Story:** As a data administrator, I want the system to match campus names despite formatting differences, so that counts are accurately mapped to the correct campuses

#### Acceptance Criteria

1. WHEN comparing campus names, THE Google Apps Script System SHALL normalize Source Sheet campus names by removing teacher name suffixes (text after " - ")
2. WHEN comparing campus names, THE Google Apps Script System SHALL convert numeric indicators in Source Sheet names (e.g., "1", "2", "3") to match the "#" format used in ALE Counts Sheet (e.g., "#1", "#2", "#3")
3. WHEN comparing campus names, THE Google Apps Script System SHALL perform case-insensitive matching between normalized names
4. WHEN a campus name from ALE Counts Sheet matches a normalized Source Sheet campus name, THE Google Apps Script System SHALL write the corresponding count value to column E of that row
5. WHEN a campus name from ALE Counts Sheet has no matching entry in any Source Sheet, THE Google Apps Script System SHALL leave column E empty for that campus

### Requirement 3

**User Story:** As a data administrator, I want the system to handle special campus naming cases, so that all campus types are correctly matched

#### Acceptance Criteria

1. WHEN processing campus names containing "AU" designation, THE Google Apps Script System SHALL preserve the "AU" designation during name normalization
2. WHEN processing multi-level campuses like "Holmgreen (MS)" or "Holmgreen (HS)", THE Google Apps Script System SHALL match based on the level indicator in parentheses
3. WHEN processing campus names with multiple numeric indicators (e.g., "Connally AU #1"), THE Google Apps Script System SHALL correctly handle both the AU designation and numeric suffix
4. WHEN processing campus names in rows D4:D113 (Elementary), D116:D160 (Middle School), D163:D206 (High School), and D208:D212 (Special Programs), THE Google Apps Script System SHALL apply matching logic to all campus ranges

### Requirement 4

**User Story:** As a data administrator, I want the count aggregation to run automatically after grade level extraction completes, so that the ALE Counts sheet is updated without additional manual steps

#### Acceptance Criteria

1. WHEN the extractESGradeLevelValue function completes execution, THE Google Apps Script System SHALL automatically trigger the campus count aggregation function
2. WHEN the extractMSGradeLevelValue function completes execution, THE Google Apps Script System SHALL automatically trigger the campus count aggregation function
3. WHEN the extractHSGradeLevelValue function completes execution, THE Google Apps Script System SHALL automatically trigger the campus count aggregation function
4. WHEN the aggregation function completes successfully after a grade level extraction, THE Google Apps Script System SHALL display a confirmation message indicating the number of campuses updated
5. WHEN the aggregation function encounters errors, THE Google Apps Script System SHALL display an error message with details about the failure

### Requirement 5

**User Story:** As a data administrator, I want the system to handle missing or invalid data gracefully, so that partial data issues do not prevent the entire aggregation from completing

#### Acceptance Criteria

1. WHEN a Source Sheet row contains an empty Count value, THE Google Apps Script System SHALL treat it as 0 for aggregation purposes
2. WHEN a Source Sheet row contains a non-numeric Count value, THE Google Apps Script System SHALL log a warning and skip that entry
3. WHEN the ALE Counts Sheet is missing or inaccessible, THE Google Apps Script System SHALL display an error message and halt execution
4. WHEN a Source Sheet (ES, MS, or HS) is missing or inaccessible, THE Google Apps Script System SHALL log a warning and continue processing remaining Source Sheets
5. WHEN multiple Source Sheet entries match the same ALE Counts campus name, THE Google Apps Script System SHALL sum the count values and write the total to column E
