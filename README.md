# Esploro Course Loader Data Transformer

This Python script processes a spreadsheet of course data (from Workday or similar systems), enriches it with researcher IDs, transforms and cleans various fields, and outputs a standardized Excel file ready for Esploro or similar research activity management systems.

## Features

- **Maps instructor names to researcher user IDs** using a lookup file.
- **Extracts and formats course, section, and attribute information** for Esploro compatibility.
- **Filters out canceled, preliminary, and certain instructional formats** (e.g., Clinical, Independent Study, Internship).
- **Handles missing and malformed data gracefully.**
- **Outputs a ready-to-load Excel file** with all required and optional fields.


## Requirements

Install the required Python packages with:

```bash
pip install pandas numpy openpyxl
```


## Input Files

1. **`WorkdayCourses.xlsx`**
The main Excel file containing course data.
**Required columns** (case-sensitive):
    - `Section Status`
    - `Instructional Format`
    - `Section`
    - `Instructors`
    - `Start Date`
    - `End Date`
    - `Course Tags`
    - `Course Subject`
    - `Course Number`
    - `Title`
    - `Enrollment Count`
    - `Academic Level`
    - `Academic Period`
    - `Delivery Mode`
2. **`researcher_lookup.xlsx`**
An Excel file mapping instructor names to their unique researcher user IDs.
**Required columns:**
    - `Name` (matches instructor names in `WorkdayCourses.xlsx`)
    - `researcherUserID` (unique identifier for each researcher)

## Output

- **`esploro_course_loader.xlsx`**
A transformed Excel file, ready for upload to Esploro or similar systems. It contains standardized activity and course fields, mapped researcher IDs, and cleaned data.


## How to Use

1. **Prepare your input files**
    - Place `WorkdayCourses.xlsx` and `researcher_lookup.xlsx` in the same directory as the script.
2. **Run the script**

```bash
python your_script_name.py
```

3. **Check your results**
    - The output file `esploro_course_loader.xlsx` will be created in the same directory.

## Field Mapping \& Transformation Details

- **Instructor Names to IDs:**
Instructor names are mapped to researcher user IDs using the lookup file. If a name is not found, the field is left blank.
- **GENR Attributes Extraction:**
Course tags containing `GENR-` codes are extracted and formatted as `activity.{code}` (e.g., `GENR-ABC` → `activity.abc`).
- **Course Section and Subject:**
    - Extracts the subject code from parentheses in the `Course Subject` field.
    - Extracts the section number from the `Section` field.
- **Course Type:**
Formats instructional format into standardized course type codes (e.g., `Independent Study` → `course.independentStudy`).
- **Date Formatting:**
Start and end dates are formatted as `MM/DD/YYYY`.
- **Filtering:**
    - Removes rows where `Section Status` is `Canceled` or `Preliminary`.
    - Removes rows with instructional formats `Clinical`, `Independent Study`, or `Internship`.
- **Additional Fields:**
Adds empty or default values for required Esploro fields not present in the input.


## Example Workflow

```bash
# Ensure required packages are installed
pip install pandas numpy openpyxl

# Place your input files in the script directory
# Run the script
python your_script_name.py

# Find your transformed file as esploro_course_loader.xlsx
```


## Troubleshooting

- **Missing Columns:**
Ensure your input files have all required columns with correct names.
- **Researcher Not Found:**
If an instructor name is missing from the lookup, the researcher ID will be blank.
- **Excel File Locked:**
Make sure output files are closed before running the script.


## License

This script is provided as-is for internal data processing and research administration purposes.
