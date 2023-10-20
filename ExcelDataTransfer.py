import os
import pandas as pd
import json
import xlwings as xw

class ExcelDataTransfer:
    def __init__(self, config_file):
        with open(config_file, 'r') as f:
            self.configs = json.load(f)

    def transfer_data(self):
        for config in self.configs:
            # Load spreadsheet
            source_wb = xw.Book(config['source_file'])

            # Iterate over all sheets in the source workbook
            for source_sheet in source_wb.sheets:
                # Load a sheet into a DataFrame by its name
                df1 = source_sheet.range('A1').options(pd.DataFrame, expand='table').value

                # Check if there is any data in the source sheet
                if df1.empty:
                    print(f"The sheet '{source_sheet.name}' in the file '{config['source_file']}' has no data.")
                    continue

                # Check if the destination file exists
                if os.path.exists(config['destination_file']):
                    destination_wb = xw.Book(config['destination_file'])
                else:
                    destination_wb = xw.Book()
                    destination_wb.save(config['destination_file'])

                # Check if the sheet exists in the destination file
                if config['destination_sheet'] in [sheet.name for sheet in destination_wb.sheets]:
                    destination_sheet = destination_wb.sheets[config['destination_sheet']]
                else:
                    destination_sheet = destination_wb.sheets.add(config['destination_sheet'])

                # Write the dataframe object into excel file
                destination_sheet.range('A1').value = df1

                print(f"Data has been written to '{config['destination_file']}'")

                # Save and close the workbook
                destination_wb.save()
                destination_wb.close()

                # Exit the loop over the sheets as soon as we find one with data
                break

            # Close the source workbook without saving
            source_wb.close()

config_file = r"C:\Arun\scripts\python\RunMacro\Application\ConfigFiles\config.json"

# Usage:
transfer = ExcelDataTransfer(config_file)
transfer.transfer_data()
