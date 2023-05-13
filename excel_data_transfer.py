import pandas as pd
import openpyxl
import sys

def main(argv):
    filename = argv[1]
    start = int(argv[2])
    end = int(argv[3])

    xlsx = pd.ExcelFile(filename, engine='openpyxl')
    df = xlsx.parse(xlsx.sheet_names[0])

    # Copy and format the columns of interest
    formatted_df = df.iloc[start:end, [1, 6, 13, 17]].copy()
    formatted_df.iloc[:, 0] = pd.to_datetime(formatted_df.iloc[:, 0], format='%d/%m/%Y').dt.date
    formatted_df.iloc[:, 2] = pd.to_numeric(formatted_df.iloc[:, 2].str.replace(',', ''), errors='coerce')
    formatted_df.iloc[:, 3] = pd.to_numeric(formatted_df.iloc[:, 3].str.replace(',', ''), errors='coerce')

    # Here we reverse the order of the rows
    formatted_df = formatted_df.iloc[::-1]

    with pd.ExcelWriter(filename, engine='openpyxl', mode='a') as writer:
        formatted_df.to_excel(writer, sheet_name='Formatted data', index=False)

if __name__ == "__main__":
    main(sys.argv)