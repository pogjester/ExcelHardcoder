import openpyxl

def create_hardcoded_copy(source_file, destination_file):
    # Load the source workbook with data_only=True to get the values
    source_wb = openpyxl.load_workbook(source_file, data_only=True)
    
    # Save the workbook directly to the destination file
    source_wb.save(destination_file)

# Example usage:
source_file = r"FILEPATH"
destination_file = r"FILEPATH"
create_hardcoded_copy(source_file, destination_file)
