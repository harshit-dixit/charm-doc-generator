import pandas as pd

# Data based on the documents and new requirements.
# Add more rows to this dictionary to generate more sets of documents.
data = {
    'Program Name': ['ZGR_INDAS_CFS_RS_CRS_FINAL'],
    'Created By': ['Harshit Dixit'],
    'Change Number': [2025007382],
    'Job Log Number': [2025007376],
    'Technical Name': ['ZGR_INDAS_CFS_RS_CRS_FINAL'],
    'Description': ['ZGR_INDAS_CFS_RS_CRS_FINAL'],
    'Test Condition': ['Changes in existing report as per user requirement.'],
    'Customer Requirement': ['New report according to user specification'],
    'Test Plan Prepared By': ['Harshit Dixit'],
    'Test Plan Reviewed By': ['Jyosyula Siva Amrutha'],
    'Testing By': ['Jyosyula Siva Amrutha'],
    'Testing Reviewed By': ['Indraneel Mazumder'],
    'Test Result Prepared By': ['Harshit Dixit'],
    'Test Result Reviewed By': ['Swagata Roy'],
    'Date': ['07.07.2025'],
    'Screenshot Path': ['C:/Users/seema/Pictures/Screenshots/Screenshots(2).png']
}

# Create DataFrame
df = pd.DataFrame(data)

# Create a Pandas Excel writer using XlsxWriter as the engine.
try:
    writer = pd.ExcelWriter('charm_data.xlsx', engine='xlsxwriter')
    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='Sheet1', index=False)
    # Close the Pandas Excel writer and output the Excel file.
    writer.close()
    print("Successfully created charm_data.xlsx")
    print("You can now edit this file to add more rows for batch processing.")
except ImportError:
    print("Error: 'xlsxwriter' is not installed. Please install it using: pip install xlsxwriter")
except Exception as e:
    print(f"An error occurred: {e}")

