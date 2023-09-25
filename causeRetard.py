# excel_data_analyzer.py
import pandas as pd
from tabulate import tabulate

class ExcelDataAnalyzer:
    def __init__(self, file_path):
        self.file_path = file_path
        self.df = None
        self.extracted_data = None
        self.vols = []

    def read_excel_file(self):
        try:
            self.df = pd.read_excel(self.file_path, header=None)
            
            # Extract data from row 6 onwards and specific columns (0, 2, and 10)
            self.extracted_data = self.df.iloc[6:, [0, 2, 10]]
            #print(self.extracted_data)
            print(tabulate(self.extracted_data, headers=['vol', 'N°', 'cause'], tablefmt='pretty'))
            
            # Iterate through rows and create dictionaries in self.vols
            for _, row in self.extracted_data.iterrows():
                #print(row[0])
                if not pd.isna(row[0]) and  not pd.isna(row[10]):  # Check if 'vol' is not 'nan'
                    self.vols.append({'vol': row[0], 'N°': (row[2]), 'cause': row[10]})
            
            # Print the list of dictionaries
            print(self.vols)
                
        except FileNotFoundError:
            print(f"File '{self.file_path}' not found. Please provide the correct file path.")
        except Exception as e:
            print(f"An error occurred: {str(e)}")
            print(e.with_traceback)
analyzer = ExcelDataAnalyzer('FR_ETAT_DES_RETARD.Vols_Depart (cause) (1).xlsx')
analyzer.read_excel_file()