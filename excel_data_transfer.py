import argparse
import locale
import openpyxl
from datetime import datetime

# Function to convert date string (DD/MM/YYYY) to date object
def convert_date(date_str):
    day, month, year = map(int, date_str.split("/"))
    return datetime(year, month, day).date()

def main():
    # Set up command line argument parser
    parser = argparse.ArgumentParser(description='Manipulate Excel Spreadsheet.')
    parser.add_argument('filename', help='Excel file name')
    parser.add_argument('start', type=int, help='Start row')
    parser.add_argument('end', type=int, help='End row')
    args = parser.parse_args()

    # Set locale to handle numeric conversions
    locale.setlocale(locale.LC_NUMERIC, '')

    # Indexes of columns 'B', 'G', 'N' and 'R'
    columns_index = [1, 6, 13, 17]
    
    try:
        # Load workbook
        wb = openpyxl.load_workbook(args.filename)
        # Set the first sheet as the active sheet
        ws = wb.active

        data = []
        
        # Iterate over specified range of rows using ws.iter_rows
        for row_cells in ws.iter_rows(min_row=args.start, max_row=args.end):
            row = []
            for column in columns_index:
                cell = row_cells[column]
                # Skip if cell has no value
                if cell.value is None:
                    continue
                # Convert date string to date object for column 'B'
                if column == 1:
                    row.append(convert_date(cell.value))
                # Convert string to float for columns 'N' and 'R'
                elif column in {13, 17}:
                    row.append(locale.atof(cell.value))
                else:
                    row.append(cell.value)
            data.append(row)
        
        # Reverse the data list so oldest entries are at the top
        data.reverse()
        
        # Remove and create the "Formatted data" sheet
        if "Formatted data" in wb.sheetnames:
            wb.remove(wb["Formatted data"])
        formatted_ws = wb.create_sheet("Formatted data")
        
        # Append the rows to the new sheet in inverse order
        for row in data:
            formatted_ws.append(row)
        
        # Save the modified workbook
        wb.save(args.filename)
    except FileNotFoundError:
        print(f"File {args.filename} not found.")
    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == "__main__":
    main()
