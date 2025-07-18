# Automated ChaRM Document Generation Prompt

**Objective**: Automate the creation of SAP ChaRM documents (`Spec.docx`, `Test Plan.docx`, `Test Results.docx`) by populating them with data from an Excel file.

**Instructions for the AI**:

Execute the following plan to provide a complete solution for the user. The user has provided three document templates (`.docx` files that will be converted to `.pdf` for your analysis) that require certain fields to be dynamically replaced.

**Your Task:**

1.  **Analyze Document Structure**: Read the provided PDF versions of the documents (`Spec.pdf`, `Test Plan.pdf`, `Test Results.pdf`) to identify all the fields that need to be replaced. This includes standard text fields and an image (screenshot). The fields to look for are:
    *   `Program Name`
    *   `Created By`
    *   `Change Number`
    *   `Job Log Number`
    *   `Technical Name`
    *   `Description`
    *   `Test Condition`
    *   `Customer Requirement`
    *   Names of personnel (e.g., `Harshit Dixit`, `Swagata Roy`) in roles like "Prepared by" or "Reviewed by".
    *   A screenshot of a report.

2.  **Create an Excel Data Source (`create_excel.py`)**:
    *   Generate a Python script named `create_excel.py`.
    *   This script will use the `pandas` and `xlsxwriter` libraries to create an Excel file named `charm_data.xlsx`.
    *   The Excel file should contain columns for all the dynamic fields identified in Step 1, including a `Screenshot Path` column.
    *   Populate the first row with the data extracted from the sample documents.

3.  **Create the Document Generator Script (`generate_charm_docs.py`)**:
    *   Generate a second Python script named `generate_charm_docs.py`.
    *   This script will use the `pandas` and `python-docx` libraries.
    *   It must read data from `charm_data.xlsx`.
    *   For each row in the Excel file, it should:
        *   Load the `.docx` templates (`Spec.docx`, `Test Plan.docx`, `Test Results.docx`).
        *   Replace placeholders (e.g., `{{ PROGRAM_NAME }}`) in the paragraphs, tables, headers, and footers of the documents.
        *   Replace a specific placeholder (e.g., `{{ SCREENSHOT_OUTPUT }}`) with the image specified in the `Screenshot Path` column.
        *   Save a new, populated `.docx` file, appending the `Change Number` to the filename to ensure uniqueness.

4.  **Create a `requirements.txt` File**:
    *   Generate a `requirements.txt` file listing all necessary Python libraries (`pandas`, `xlsxwriter`, `python-docx`).

5.  **Create a `README.md` File**:
    *   Generate a comprehensive `README.md` file that explains:
        *   The project's purpose.
        *   Step-by-step instructions on how to set up and run the scripts.
        *   How to install dependencies using `requirements.txt`.
        *   A clear table of all the placeholders that must be inserted into the `.docx` templates.
        *   A to-do list for the user to follow.

---

## Implementation Overview

This section outlines the technical implementation of the automated document generation system.

### 1. Dependencies

The following Python libraries are required. They can be installed via the `requirements.txt` file.

```bash
pandas
xlsxwriter
python-docx
```

### 2. Data Source: Excel File

An Excel file (`charm_data.xlsx`) serves as the single source of truth for all dynamic data. Each row corresponds to one set of ChaRM documents.

**Example Structure:**

| Program Name | Created By | Change Number | ... | Screenshot Path      |
| :----------- | :--------- | :------------ | :-- | :------------------- |
| Report_ABC   | John Doe   | CHG001        | ... | C:\path\to\image.png |

### 3. Word Document Templates

The `.docx` templates (`Spec.docx`, `Test Plan.docx`, `Test Results.docx`) must be prepared with specific placeholders where data will be inserted.

**Placeholder Format**: `{{ PLACEHOLDER_NAME }}`

**Example Placeholders**:
- `{{ PROGRAM_NAME }}`
- `{{ CREATED_BY }}`
- `{{ SCREENSHOT_OUTPUT }}`

### 4. Core Python Scripts

Two Python scripts manage the workflow:

#### `create_excel.py`
- **Purpose**: To generate a template `charm_data.xlsx` file with the correct headers and one row of sample data.
- **Libraries**: `pandas`, `xlsxwriter`

#### `generate_charm_docs.py`
- **Purpose**: The main script that performs the automation.
- **Logic**:
    1. Reads the `charm_data.xlsx` file into a pandas DataFrame.
    2. Iterates through each row of the DataFrame.
    3. For each row, it opens the three `.docx` templates.
    4. It systematically searches for and replaces all defined placeholders in paragraphs, tables, headers, and footers with the corresponding data from the row.
    5. It replaces the screenshot placeholder with the image file specified in the `Screenshot Path`.
    6. It saves the modified documents with a unique name incorporating the Change Number (e.g., `Spec_Output_CHG001.docx`).
- **Libraries**: `pandas`, `python-docx`

### 5. Final Output

The process results in a set of populated `.docx` files for each entry in the Excel sheet, ready for review and use in the ChaRM process.