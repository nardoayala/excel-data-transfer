#!/usr/bin/python3
import argparse
import openpyxl
import pyperclip

def format_date(date_str):
    date = date_str.split("/")
    day = date[0]
    month = date[1]
    year = date[2]
    return f"{month}/{day}/{year}"

def format_number(string):
    return string.replace(",", "")


def main():
    parser = argparse.ArgumentParser(
        description='Manipulate Excel Spreadsheet.'
    )
    parser.add_argument('filename', help='Excel file name')
    parser.add_argument('start', type=int, help='Start row')
    parser.add_argument('end', type=int, help='End row')
    args = parser.parse_args()

    # Indexes of columns 'B', 'G', 'N'
    columns_index = [1, 6, 13]

    try:
        # Load workbook
        wb = openpyxl.load_workbook(args.filename)
        # Set the first sheet as the active sheet
        ws = wb.worksheets[0]

        data = []

        for row_cells in ws.iter_rows(min_row=args.start, max_row=args.end):
            row = []
            for column in columns_index:
                cell = row_cells[column]
                # Skip if cell has no value
                if cell.value is None:
                    continue
                if column == 1:
                    row.append(format_date(cell.value))
                    # Insert an empty column after 'B'
                    row.append("")
                elif column == 13:
                    row.append(format_number(cell.value))
                    # Append a column with the string "Facebank" after "N"
                    row.append("Facebank")
                else:
                    row.append(cell.value)
            if "COMISION" in row[2]:
                row[1] = "Bank fees"
            data.append(row)

        # Reverse the data list so oldest entries are at the top
        data.reverse()
        
        data_for_clipboard = ""
        for row in data:
            if row != data[-1]:
                data_for_clipboard += "\t".join(row) + "\n"
            else:
                data_for_clipboard += "\t".join(row)
        pyperclip.copy(data_for_clipboard)

    except FileNotFoundError:
        print(f"File {args.filename} not found.")
    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":
    main()
