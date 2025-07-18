# ChaRM Document Automation

This project automates the creation of three essential SAP ChaRM documents: the Specification, Test Plan, and Test Results. It uses a Python script to read data from an Excel spreadsheet and populate placeholder fields in template Word documents, significantly speeding up the documentation process.

## Features

- **Batch Processing**: Generate multiple sets of documents from a single Excel file.
- **Preserves Formatting**: Retains the original styling of your Word templates.
- **Dynamic Placeholders**: Replaces text in paragraphs, tables, headers, and footers.
- **Image Replacement**: Automatically inserts screenshots into the documents.
- **Data Consistency**: Enforces consistent data across related fields.

## Prerequisites

- Python 3.6 or higher
- Microsoft Word (.docx) templates

## How It Works

1.  **`charm_data.xlsx`**: An Excel file where you store the data for each document set. Each row represents a new set of documents.
2.  **`.docx` Templates**: Your `Spec.docx`, `Test Plan.docx`, and `Test Results.docx` files must be updated to include placeholders (e.g., `{{ PROGRAM_NAME }}`).
3.  **`generate_charm_docs.py`**: The main Python script that reads the Excel data and populates the Word templates.

## Setup and Usage

Follow these steps to set up and run the automation:

### Step 1: Install Dependencies

Install all the required Python libraries using the `requirements.txt` file:

```bash
pip install -r requirements.txt
```

### Step 2: Prepare Your Data

1.  A `charm_data.xlsx` file will be created for you if it doesn't exist. If you need to create it manually, you can run the `create_excel.py` script:

    ```bash
    python create_excel.py
    ```

2.  Open `charm_data.xlsx` and fill in the required information. Each row will produce a new set of documents. The script will automatically enforce the following relationships:
    - **Description** will be the same as **Program Name**.
    - **Test Plan Prepared By** will be the same as **Created By**.
    - **Testing By** will be the same as **Test Plan Reviewed By**.
    - **Test Result Prepared By** will be the same as **Created By**.

3.  **Crucially**, update the `Screenshot Path` column with the absolute file path to your screenshot image for each row.

### Step 3: Prepare Your Word Templates

For the script to work, you must replace the dynamic text in your `Spec.docx`, `Test Plan.docx`, and `Test Results.docx` templates with the following placeholders.

**Important**: The placeholders must be typed exactly as they appear below, including the curly braces and spacing.

| Field in Excel          | Placeholder in Word Document        |
| ----------------------- | ----------------------------------- |
| `Program Name`          | `{{ PROGRAM_NAME }}`                |
| `Created By`            | `{{ CREATED_BY }}`                  |
| `Change Number`         | `{{ CHANGE_NUMBER }}`               |
| `Job Log Number`        | `{{ JOB_LOG_NUMBER }}`              |
| `Technical Name`        | `{{ TECHNICAL_NAME }}`              |
| `Description`           | `{{ DESCRIPTION }}`                 |
| `Test Condition`        | `{{ TEST_CONDITION }}`              |
| `Customer Requirement`  | `{{ CUSTOMER_REQUIREMENT }}`        |
| `Test Plan Prepared By` | `{{ TEST_PLAN_PREPARED_BY }}`       |
| `Test Plan Reviewed By` | `{{ TEST_PLAN_REVIEWED_BY }}`       |
| `Testing By`            | `{{ TESTING_BY }}`                  |
| `Testing Reviewed By`   | `{{ TESTING_REVIEWED_BY }}`         |
| `Test Result Prepared By`| `{{ TEST_RESULT_PREPARED_BY }}`     |
| `Test Result Reviewed By`| `{{ TEST_RESULT_REVIEWED_BY }}`     |
| `Screenshot Path`       | `{{ SCREENSHOT_OUTPUT }}`           |

### Step 4: Run the Generator

Execute the main script to generate your documents:

```bash
python generate_charm_docs.py
```

The script will create a new set of `.docx` files for each row in your Excel sheet. The output files will be named using the convention `[TemplateName]_[ChangeNumber].docx` (e.g., `Spec_Output_2025007382.docx`).

## To-Do

- [ ] Update `Spec.docx` with placeholders.
- [ ] Update `Test Plan.docx` with placeholders.
- [ ] Update `Test Results.docx` with placeholders.
- [ ] Populate `charm_data.xlsx` with your data.
- [ ] Run the script to generate the final documents.
