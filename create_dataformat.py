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
                year_title = (
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
                site_data[site_name]["1. Uitleg en Bedrijfsgegevens"] = (
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


def update_sheet(excel_template_data_sheet, excel_template_layout_sheet):
    sheet_rows = dataframe_to_rows(excel_template_data_sheet, index=False, header=False)
    for r_idx, row in enumerate(sheet_rows, 1):
        for c_idx, value in enumerate(row, 1):
            cell = excel_template_layout_sheet.cell(row=r_idx, column=c_idx)
            cell_coordinate = f"{get_column_letter(c_idx)}{r_idx}"
            if isinstance(cell, openpyxl.cell.MergedCell):
                # Find the top-left cell of the merged range
                for merged_range in excel_template_layout_sheet.merged_cells.ranges:
                    if cell_coordinate in merged_range:
                        top_left_cell = excel_template_layout_sheet.cell(
                            row=merged_range.min_row, column=merged_range.min_col
                        )
                        top_left_cell.value = value
                        break
            else:
                cell.value = value


def insert_site_data():
    print("Inserting site data")
    for site_name in site_data:
        # excel_template_layout = openpyxl.load_workbook(SOURCE_FILE)
        excel_template_layout = xw.Book(SOURCE_FILE)

        for sheet_name in excel_template_layout.sheet_names:
            excel_template_layout.sheets[sheet_name]["C15"].value = site_name
            # excel_template_layout_sheet = excel_template_layout[sheet_name]
            # excel_template_data_sheet = pd.read_excel(
            #     SOURCE_FILE, engine="openpyxl", sheet_name=sheet_name
            # )

            # if sheet_name == "1. Uitleg en Bedrijfsgegevens":
            #     data = site_data[site_name][sheet_name]
            #     print(data)
            #     excel_template_data_sheet.iloc[11, 2] = data["latitude"]
            #     excel_template_data_sheet.iloc[12, 2] = data["longitude"]
            #     excel_template_data_sheet.iloc[13, 2] = site_name

            # update_sheet(excel_template_data_sheet, excel_template_layout_sheet)

        # Create output file
        os.makedirs(output_folder, exist_ok=True)
        excel_template_layout.save(f"{output_folder}/{site_name}.xlsx")
        excel_template_layout.close()
        exit(0)


# Main
def main():
    # Extract data from the input file
    extract_site_data()

    # Create output files
    # print(f"Creating output files")
    # for site_name in site_data:
    #     create_output_file(site_name)

    # Put data in output files
    insert_site_data()


if __name__ == "__main__":
    main()
