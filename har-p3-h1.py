import logging
import datetime
import pandas as pd
from pandas import ExcelWriter
from shutil import copyfile

# Päivä 3
# Harjoitus 1
#
# Käsitellään Northwind tietokannasta Excel-tiedostoon vietyjä tuotetietoja (products.xlsx).
# Luetaan tuotetietoja Excel-tiedostosta ja tallennetaan ne yhden Excel-tiedoston kolmelle eri välilehdelle.
# Näin muodostetaan products_divided.xlsx tiedosto, jossa on kolme välilehteä:
# - alive (kaikki tuotteet, joiden discontinued-sarake on False)
# - alive_50k (tuotteet joiden varastoarvo on yli 1000 rahayksikköä; varastoarvo = unitsinstock * unitprice)
# - discontinued (kaikki tuotteet, joiden discontinued-sarake on True)
#

harjoitus_name = "p3-h1"

input_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\harjoitus_data\\products.xlsx"
output_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\harjoitus_data\\products_divided.xlsx"
empty_excel_file = r"D:\\CAREERIA\\it-opiskelu\\2025-syksy\\robot\\harjoitus_data\\empty.xlsx"

# Creating an empty output file to avoid "At least one sheet must be visible" error from pandas
# when trying to write to a non-existing file. So we copy an empty Excel file to the output file path.
copyfile(empty_excel_file, output_file)

program_start_time = datetime.datetime.now()

# Set up logging file 
log_filename = harjoitus_name + "_log.txt"
logging.basicConfig(
    filename=log_filename,
    level=logging.INFO,
    format="%(asctime)s %(message)s",
    encoding="utf-8",
    filemode="w",
)  # Overwrite log file each run


def write_to_log(s):
    logging.info(s)
    

def rpa_excel_main():
    write_to_log(f"Aloitusaika: {program_start_time.strftime('%d.%m.%Y %H:%M:%S')}")
   
    try:
        # Read the input Excel file
        df = pd.read_excel(input_file)
        write_to_log(f"Tuotetiedot luettu tiedostosta {input_file}.")

        # sort the DataFrame by ProductName
        df.sort_values(by='ProductName', inplace=True)

        # Create a Pandas Excel writer using openpyxl as the engine
        with ExcelWriter(output_file, mode="a", engine='openpyxl') as writer:
            # Write alive products (Discontinued == False)
            df_alive = df[~df['Discontinued']]  # ~ is the NOT operator
            df_alive.to_excel(writer, sheet_name='alive', index=False)

            # Write alive products with stock value over 1000
            df_alive_50k = df_alive[df_alive['UnitsInStock'] * df_alive['UnitPrice'] > 1000]
            df_alive_50k.to_excel(writer, sheet_name='alive_50k', index=False)

            # Write discontinued products (Discontinued == True)
            df_discontinued = df[df['Discontinued']]
            df_discontinued.to_excel(writer, sheet_name='discontinued', index=False)

            # delete the default sheet created by default in the empty Excel file
            if 'Sheet1' in writer.book.sheetnames:
                std = writer.book['Sheet1']
                writer.book.remove(std)

        write_to_log(f"Tuotetiedot lajiteltu ja tallennettu tiedostoon {output_file}.")

    except Exception as e:
        write_to_log(f"Virhe: {e}")

    program_end_time = datetime.datetime.now()
    program_duration = program_end_time - program_start_time
    write_to_log(f"Lopetusaika: {program_end_time.strftime('%d.%m.%Y %H:%M:%S')}")
    write_to_log(f"Ajon kesto: {str(program_duration)}")
    write_to_log("****** DONE ******")

if __name__ == "__main__":
    rpa_excel_main()
    print("For more details, see the log file " + log_filename)
    print("****** DONE ******")
