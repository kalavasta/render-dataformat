import sys
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.utils import get_column_letter
import xlwings as xw
import os
import pandas as pd
import re
import shutil
from functions import strip_string, create_file

SOURCE_FILE = "templates/Dataformat template.xlsx"
INPUT_DATA_SHEETS = ["Chemie", "Raffinage", "Voedsel", "OpslagenVervoer"]

input_file = sys.argv[1]
output_folder = sys.argv[2]
site_data = {}


def create_output_file(site_name):
    os.makedirs(output_folder, exist_ok=True)

    destination_file = f"{output_folder}/{strip_string(site_name)}.xlsx"
    shutil.copy(SOURCE_FILE, destination_file)
    print(f"Created {destination_file}")


def get_company_info(site_name):
    excel_content = pd.read_excel(input_file, engine="openpyxl", sheet_name="plants")
    excel_content = excel_content.fillna("")

    n_rows = len(excel_content)
    site_name_found = False

    for row_n in range(0, n_rows):
        current_site_name = excel_content.iloc[row_n, 0]

        if current_site_name != site_name:
            continue

        site_name_found = True

        return {
            "site": site_name,
            "address": excel_content.iloc[row_n, 5].strip(),
            "city": excel_content.iloc[row_n, 6].strip(),
            "postal_code": excel_content.iloc[row_n, 4].strip(),
            "new_or_existing": excel_content.iloc[row_n, 7].strip(),
            "latitude": excel_content.iloc[row_n, 9],
            "longitude": excel_content.iloc[row_n, 10],
            "sector": excel_content.iloc[row_n, 8],
            "cluster": excel_content.iloc[row_n, 3].strip(),
            "ean_electricity": excel_content.iloc[row_n, 11],
            "ean_gas": excel_content.iloc[row_n, 12],
            "grid_operator_electricity": excel_content.iloc[row_n, 13].strip(),
            "grid_operator_gas": excel_content.iloc[row_n, 14].strip(),
        }

    if not site_name_found:
        print(f"Site name {site_name} not found in the plants sheet")
        sys.exit(1)


def extract_site_data():
    print("Extracting site data")

    for sheet in INPUT_DATA_SHEETS:
        excel_content = pd.read_excel(input_file, engine="openpyxl", sheet_name=sheet)
        excel_content = excel_content.fillna("")

        n_rows = len(excel_content)
        site_name = ""
        year_name = ""
        year_number = ""
        year_title = ""

        for row_n in range(3, n_rows):
            site_name_temp = excel_content.iloc[row_n, 2]
            year_name_temp = str(excel_content.iloc[row_n, 4])

            if year_name_temp != "":
                year_name = year_name_temp
                year_number = re.findall(r"\d{4}", year_name)[0]
                year_title = strip_string(
                    "base"
                    if year_number == "2021"
                    else re.split(r"\d{4} ", year_name)[1]
                )

            if (
                site_name_temp != ""
                and site_name_temp != 2030
                and site_name_temp != 2035
                and site_name_temp != 2040
                and site_name_temp != 2050
            ):
                site_name = excel_content.iloc[row_n, 2]
                site_data[site_name] = {}
                site_data[site_name][strip_string("1. Uitleg en Bedrijfsgegevens")] = (
                    get_company_info(site_name)
                )

            if site_name == "":
                continue

            if not site_data[site_name].get(year_title):
                site_data[site_name][year_title] = {}

            if not site_data[site_name][year_title].get(year_number):
                site_data[site_name][year_title][year_number] = {}

            demand_supply = excel_content.iloc[row_n, 5]

            site_data[site_name][year_title][year_number][demand_supply] = {
                "co2_fossil": excel_content.iloc[row_n, 6],
                "co2_bio": excel_content.iloc[row_n, 7],
                "electricity_anual": excel_content.iloc[row_n, 8],
                "electricity_peak": excel_content.iloc[row_n, 9],
                "natural_gas": excel_content.iloc[row_n, 10],
                "hydrogen_more_than_98": excel_content.iloc[row_n, 11],
                "hydrogen_less_than_98": excel_content.iloc[row_n, 12],
                "heat_less_than_100": excel_content.iloc[row_n, 15],
                "heat_more_than_100": excel_content.iloc[row_n, 16],
            }

    filename = "./test.json"
    create_file(filename, site_data)  # @TODO: Remove line


def insert_site_data():
    print("Inserting site data")
    for site_name in site_data:
        excel_template_layout = xw.Book(SOURCE_FILE)

        for sheet_name in excel_template_layout.sheet_names:
            print(sheet_name)
            sheet = excel_template_layout.sheets[sheet_name]
            stripped_sheet_name = strip_string(sheet_name)
            sheet_data = site_data[site_name].get(stripped_sheet_name, False)

            if sheet_data == False:
                continue

            if stripped_sheet_name == strip_string("1. Uitleg en Bedrijfsgegevens"):

                sheet["C9"].value = sheet_data.get("address", "")
                sheet["C10"].value = sheet_data.get("city", "")
                sheet["C11"].value = sheet_data.get("postal_code", "")
                sheet["C12"].value = sheet_data.get("new_or_existing", "")
                sheet["C13"].value = sheet_data.get("latitude", "")
                sheet["C14"].value = sheet_data.get("longitude", "")
                sheet["C15"].value = site_name
                sheet["C16"].value = sheet_data.get("sector", "")
                sheet["C17"].value = sheet_data.get("cluster", "")
                sheet["C18"].value = f"{sheet_data.get("ean_electricity", "")}" # @TODO: Fix rounding error
                sheet["C19"].value = f"{sheet_data.get("ean_gas", "")}" # @TODO: Fix rounding error
                sheet["C20"].value = sheet_data.get("grid_operator_electricity", "")
                sheet["C21"].value = sheet_data.get("grid_operator_gas", "")

            if (
                stripped_sheet_name == strip_string("Koersvaste Ambitie")
                or stripped_sheet_name == strip_string("Eigen Vermogen")
                or stripped_sheet_name == strip_string("Gemeenschappelijke balans")
                or stripped_sheet_name == strip_string("Horizon aanvoer")
            ):
              for year in (['2021'] + list(sheet_data.keys())):
                rows = {
                    '2021': 4,
                    '2030': 10,
                    '2035': 13,
                    '2040': 16,
                    '2050': 19,
                }

                data = site_data[site_name]['base']['2021'] if year == '2021' else sheet_data[year]

                for i in range(0, 2):
                    demand_or_supply = "Demand" if i == 0 else "Supply"
                    row = rows[year] + i
                    sheet[f"K{row}"].value = data[demand_or_supply].get("co2_fossil", "")
                    sheet[f"L{row}"].value = data[demand_or_supply].get("co2_bio", "")
                    sheet[f"M{row}"].value = data[demand_or_supply].get("electricity_anual", "")
                    sheet[f"N{row}"].value = data[demand_or_supply].get("electricity_peak", "")
                    sheet[f"P{row}"].value = data[demand_or_supply].get("natural_gas", "")
                    sheet[f"Q{row}"].value = data[demand_or_supply].get("hydrogen_more_than_98", "")
                    sheet[f"R{row}"].value = data[demand_or_supply].get("hydrogen_less_than_98", "")
                    sheet[f"S{row}"].value = data[demand_or_supply].get("heat_less_than_100", "")
                    sheet[f"T{row}"].value = data[demand_or_supply].get("heat_more_than_100", "")


        # Create output file
        os.makedirs(output_folder, exist_ok=True)
        excel_template_layout.save(f"{output_folder}/{site_name}.xlsx")
        excel_template_layout.close()
        # exit(0)


# Main
def main():
    # Extract data from the input file
    extract_site_data()

    # Put data in output files
    insert_site_data()


if __name__ == "__main__":
    main()
