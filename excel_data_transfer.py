from datetime import datetime
import openpyxl
import sys

def main(argv):
    filename = argv[1]
    start = int(argv[2])
    # Add 1 to the end row to include it in the range
    end = int(argv[3]) + 1

    wb = openpyxl.load_workbook(filename)
    # Store data from the start to the end rows inside the B, G, N and R excel columns
    columns_index = [2, 7, 14, 18]
    data = []

    # Set first sheet as active sheet
    wb.active = 0

    # Create new sheet called "Formatted data" if it doesn't exist
    if "Formatted data" not in wb.sheetnames:
        wb.create_sheet("Formatted data")

    for i in range(start, end):
        row = []
        for j in columns_index:
            ## If column is 2, convert date string (DD/MM/YYYY) to date object
            if j == 2:
                date_list = wb.active.cell(row=i, column=j).value.split("/")
                year = int(date_list[2])
                month = int(date_list[1])
                day = int(date_list[0])
                date = datetime(year, month, day).date()
                # append as excel date object
                row.append(date)
            ## If colum is 14 or 18, convert number stored as string with ',' as thousand separator to float
            elif j == 14 or j == 18:
                row.append(float(wb.active.cell(row=i, column=j).value.replace(",", "")))
            else:
                row.append(wb.active.cell(row=i, column=j).value)
        data.append(row)
    
    # Invert data list so oldest entries are at the top
    data.reverse()

    # If the new sheet already has data, delete it
    if wb["Formatted data"].max_row > 1:
        wb.remove(wb["Formatted data"])
        wb.create_sheet("Formatted data")

    # Append the rows to the new sheet in inverse order
    for row in data:
        wb["Formatted data"].append(row)

    # Save the new sheet
    wb.save(filename)

if __name__ == "__main__":
    main(sys.argv)