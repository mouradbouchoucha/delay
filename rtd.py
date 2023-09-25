import pandas as pd
import os
import csv

# Read the Excel file and save it as a CSV with the same name
excel_file = "ETAT_DES_RETARD.Vols_Depart - 2023-09-21T214932.218.xls"
df = pd.read_excel(excel_file)
base_name = os.path.splitext(os.path.basename(excel_file))[0]
csv_file = f"{base_name}.csv"
df.to_csv(csv_file, index=False)

# Read the CSV file
input_file = "/ETAT_DES_RETARD.Vols_Depart - 2023-09-21.csv"
output_file = "FR_ETAT_DES_RETARD.Vols_Depart.csv"

# Delete rows that end with 'CHA', 'REG', 'SUP'
with open(input_file, mode='r', newline='') as infile, open(output_file, mode='w', newline='') as outfile:
    reader = csv.reader(infile)
    writer = csv.writer(outfile)

    for row in reader:
        if not row or row[-1] in ('CHA', 'REG', 'SUP'):
            writer.writerow(row)
            print(row[10])

# Reopen the modified file and count the rows
with open(output_file, mode='r', newline='') as infile:
    reader = csv.reader(infile)
    tar_count = 0
    lbt_count = 0
    tux_count = 0
    autre_count = 0
    total = 0

    for row in reader:
        total += 1
        if row and row[0] == 'TAR':
            tar_count += 1
        elif row and row[0] == 'LBT':
            lbt_count += 1
        elif row and row[0] == 'TUX':
            tux_count += 1
        else:
            autre_count += 1

print(f"Total 'TAR' rows: {tar_count}")
print(f"Total 'LBT' rows: {lbt_count}")
print(f"Total 'TUX' rows: {tux_count}")
print(f"Total 'AUTRE' rows: {autre_count}")
print(f"Total rows: {total}")


