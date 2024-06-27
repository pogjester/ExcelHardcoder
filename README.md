
# Create Hardcoded Copy of Excel File

This script uses the `openpyxl` library to create a hardcoded copy of an Excel file. When you load an Excel file with `data_only=True`, it reads the values from the cells instead of any formulas. The script then saves this data to a new Excel file, effectively creating a version of the file with only the hardcoded values.

## Prerequisites

Make sure you have `openpyxl` installed. You can install it using pip if you haven't already:

```bash
pip install openpyxl
```

## Usage

1. **Download the Script**

   Save the following script to a file, for example, `copy_excel.py`.

   ```python
   import openpyxl

   def create_hardcoded_copy(source_file, destination_file):
       # Load the source workbook with data_only=True to get the values
       source_wb = openpyxl.load_workbook(source_file, data_only=True)
       
       # Save the workbook directly to the destination file
       source_wb.save(destination_file)

   # Example usage:
   source_file = r"C:\Path\To\Your\SourceFile.xlsx"
   destination_file = r"C:\Path\To\Your\DestinationFile.xlsx"
   create_hardcoded_copy(source_file, destination_file)
   ```

2. **Edit the Script**

   Update the `source_file` and `destination_file` variables with the paths to your source and destination files.

3. **Run the Script**

   Run the script using Python:

   ```bash
   python copy_excel.py
   ```

## Example

If your source file is located at `C:\Users\YourUsername\Documents\source.xlsx` and you want to save the hardcoded copy as `C:\Users\YourUsername\Documents\destination.xlsx`, you would set the `source_file` and `destination_file` variables as follows:

```python
source_file = r"C:\Users\YourUsername\Documents\source.xlsx"
destination_file = r"C:\Users\YourUsername\Documents\destination.xlsx"
```

Then, run the script:

```bash
python copy_excel.py
```

## Explanation

- **Loading the Workbook**: The script loads the source Excel workbook with `data_only=True`, which means it reads the values from the cells instead of the formulas.
- **Saving the Workbook**: The script then saves the workbook to the specified destination file. The saved file contains only the values from the original file, with all formulas replaced by their evaluated results.

This method provides a simple and efficient way to create a hardcoded copy of an Excel file, preserving all the data without any formulas.
