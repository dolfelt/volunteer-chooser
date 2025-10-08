# Volunteer Chooser

A Go-based tool for assigning volunteers to classroom parties and field trips. The tool reads volunteer sign-up data from an Excel spreadsheet and automatically assigns volunteers to events based on their preferences, ensuring fair, and random, distribution across teachers and events.

## Overview

Volunteer Chooser processes volunteer sign-ups and:
- Assigns volunteers to parties and field trips based on their selections
- Ensures each teacher gets the configured number of volunteers per event
- Prevents volunteers from being assigned to multiple parties or multiple field trips (but they can do one party and one field trip)
- Automatically assigns 2 alternates per teacher per event
- Outputs formatted Excel sheets, one per event, with all assignments

## Installation

```bash
go build -o volunteer-chooser main.go
```

## Usage

```bash
./volunteer-chooser [options]
```

### Command-line Options

- `-input` - Input XLSX file (default: `input.xlsx`)
- `-output` - Output XLSX file (default: `output.xlsx`)
- `-seed` - Random seed for reproducible assignments (default: `42`)

### Examples

```bash
# Use default settings
./volunteer-chooser

# Specify custom input and output files
./volunteer-chooser -input volunteers.xlsx -output assignments.xlsx

# Use a different random seed for different volunteer assignments
./volunteer-chooser -seed 12345
```

## Input File Format

The input Excel file must contain two sheets:

### 1. Form Responses 1 Sheet

This sheet contains volunteer sign-up data from a Google Form (or similar). The tool automatically detects column names containing these keywords (case-insensitive):

| Column Purpose | Detected Keywords | Example Values |
|----------------|-------------------|----------------|
| Teacher name | "teacher" | "Ms. Smith", "Mr. Johnson" |
| Volunteer name | "first and last name" OR "first name" + "last name" | "John Doe" OR "John" + "Doe" |
| Phone number | "phone" | "555-123-4567" |
| Email address | "email" | "john.doe@example.com" |
| Party selections | "party or parties" | "Halloween Party, Valentine's Party" |
| Field trip selections | "field trip(s)" | "Zoo Trip, Museum Visit" |

**Note**: The exact column headers don't need to match exactly - the tool searches for these keywords within the header text.

### 2. Variables Sheet

This Variable sheet defines the parties and field trips with their volunteer requirements:

#### Party Configuration (Columns A-B)

| Column A | Column B |
|----------|----------|
| Party Name | Count Per Teacher |
| Halloween Party | 3 |
| Valentine's Party | 2 |
| End of Year Party | 4 |

- **Party Name** (Column A): Name of the party (must match text in Form Responses)
- **Count Per Teacher** (Column B): Number of volunteers needed for each teacher

#### Field Trip Configuration (Columns C-E)

| Column C | Column D | Column E |
|----------|----------|----------|
| Field Trip Name | Teachers | Count |
| Zoo Trip | Ms. Smith\|Mr. Johnson | 5 |
| Museum Visit | ALL | 3 |
| Science Center | Ms. Brown | 4 |

- **Field Trip Name** (Column C): Name of the field trip (must match text in Form Responses)
- **Teachers** (Column D): Pipe-separated list of teacher names, or "ALL" for all teachers
- **Count** (Column E): Number of volunteers needed per applicable teacher

**Teacher Specification Examples**:
- `ALL` - Field trip applies to all teachers found in form responses
- `Ms. Smith|Mr. Johnson` - Field trip only for these specific teachers
- `Ms. Brown` - Field trip for a single teacher

## Output Format

The tool generates an Excel file with one sheet per event (parties first, then field trips in the order defined in Variables sheet).

Each sheet contains:
- Event name as the title
- Grouped sections by teacher
- Volunteer details: Name, Phone, Email
- Clear visual separation between primary assignments and alternates

### Assignment Rules

1. **Primary Assignments**:
   - Each volunteer can be assigned to at most one party and one field trip
   - Volunteers are randomly shuffled (using the seed) for fair distribution
   - Assignments are made until the configured count is reached for each teacher

2. **Alternates**:
   - 2 alternates are assigned per teacher per event
   - Alternates are marked with italic text and shaded background
   - Alternates are not used as primary volunteers for any other event
   - If insufficient unique volunteers exist, the tool may reuse volunteers who weren't selected for other events

## Example Workflow

1. **Create input file** (`volunteers.xlsx`):
   - "Form Responses 1" sheet with volunteer data
   - "Variables" sheet with party and field trip configurations

2. **Run the tool**:
   ```bash
   ./volunteer-chooser -input volunteers.xlsx -output assignments.xlsx -seed 100
   ```

3. **Review output** (`assignments.xlsx`):
   - Each event has its own sheet
   - Volunteers are organized by teacher
   - Primary volunteers listed first, followed by alternates

4. **Adjust if needed**:
   - Change the seed value to generate a different random assignment
   - Modify Variables sheet to change volunteer counts
   - Re-run the tool

## Notes

- The tool uses deterministic randomization based on the seed, so running with the same seed produces identical results
- Teacher names must match exactly between Form Responses and Variables sheets
- Event names must match exactly between Form Responses selections and Variables sheet definitions
- Sheet names in the output are sanitized to comply with Excel's 31-character limit and character restrictions

## Dependencies

- [github.com/xuri/excelize/v2](https://github.com/xuri/excelize) - Excel file handling

Install dependencies with:
```bash
go get github.com/xuri/excelize/v2
```
