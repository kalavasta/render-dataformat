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
import time
import math

SOURCE_FILE = "templates/Dataformat template.xlsx"
INPUT_DATA_SHEETS = ["Chemie", "Raffinage", "Voedsel", "OpslagenVervoer"]
FLEX_MEANS = {
    "Flexible production of heat": "Hybride warmte",
    "Flexible use of CHP": "Flexibele inzet van WKK",
    "Buffering of electricity": "Opslag van elektriciteit",
    "Buffering of heat": "Thermische buffer",
    "Flexibility in process": "Procesflexibiliteit",
    "Electrochemistry": "Electrochemie (inclusief groene H2 productie)",
}
AVAILABILITY = {"Medium": "medium", "High": "hoog", "Low": "laag"}
STORY_LINES = [
    "Koersvaste ambitie",
    "Eigen vermogen",
    "Gemeenschappelijke balans",
    "Horizon aanvoer",
]
REDUCTION_SHIFT = {
    "Reduction": "Reductie",
    "Shift": "Verschuiving",
}

story_lines_stripped = []
for item in STORY_LINES:
    story_lines_stripped = story_lines_stripped + [strip_string(item)]

input_file = sys.argv[1]
output_folder = sys.argv[2]
site_data = {}


def get_company_info(site_name, sheet_name, parent_row_n):
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
            "ean_electricity": f"'{excel_content.iloc[row_n, 11]}",
            "ean_gas": f"'{excel_content.iloc[row_n, 12]}",
            "grid_operator_electricity": excel_content.iloc[row_n, 13].strip(),
            "grid_operator_gas": excel_content.iloc[row_n, 14].strip(),
        }

    if not site_name_found:
        exit(
            f"Error: Site name `{site_name}` (sheet `{sheet_name}`, row {parent_row_n + 2}) not found in the plants sheet"
        )


def extract_site_data():
    print("> Extracting site data")

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
                    get_company_info(site_name, sheet, row_n)
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

    excel_content = pd.read_excel(input_file, engine="openpyxl", sheet_name="flex")
    excel_content = excel_content.fillna("")
    n_rows = len(excel_content)
    SHEET_NAME = "4. Flexibiliteit (begeleid)"
    stripped_sheet_name = strip_string(SHEET_NAME)

    for row_n in range(0, n_rows):
        site_name = excel_content.iloc[row_n, 1]

        if not site_data.get(site_name):
            continue

        year_number = str(excel_content.iloc[row_n, 4])
        year_titles = (
            "base"
            if excel_content.iloc[row_n, 5] == ""
            else excel_content.iloc[row_n, 5]
        )

        if year_number == "2021" and year_titles != "base":
            exit(
                f"Error: Storyline for `{site_name}` year 2021 (sheet `flex`, row {row_n + 2}) should be empty"
            )

        year_titles_list = re.split(", ", year_titles)

        if not site_data[site_name].get(stripped_sheet_name):
            site_data[site_name][stripped_sheet_name] = {}

        if not site_data[site_name][stripped_sheet_name].get(year_number):
            site_data[site_name][stripped_sheet_name][year_number] = {}

        for year_title in year_titles_list:
            if not strip_string(year_title) in (story_lines_stripped + ["base"]):
                exit(
                    f"Error: `{year_title}` (sheet `flex`, row {row_n + 2}) is not a valid storyline name, please use one of the following: {STORY_LINES}"
                )

            if not site_data[site_name][stripped_sheet_name][year_number].get(
                strip_string(year_title)
            ):
                site_data[site_name][stripped_sheet_name][year_number][
                    strip_string(year_title)
                ] = []

            site_data[site_name][stripped_sheet_name][year_number][
                strip_string(year_title)
            ] = site_data[site_name][stripped_sheet_name][year_number][
                strip_string(year_title)
            ] + [
                {
                    "flex_means": FLEX_MEANS[excel_content.iloc[row_n, 6]],
                    "flexible_power": excel_content.iloc[row_n, 7],
                    "availability": AVAILABILITY[excel_content.iloc[row_n, 8]],
                    "operational_hours": excel_content.iloc[row_n, 9],
                    "buffer_capacity": excel_content.iloc[row_n, 10],
                    "performance_coefficient": excel_content.iloc[row_n, 11],
                    "reduction_shift": (
                        ""
                        if excel_content.iloc[row_n, 12] == ""
                        else REDUCTION_SHIFT[excel_content.iloc[row_n, 12]]
                    ),
                    "consecutive_hours": excel_content.iloc[row_n, 13],
                }
            ]


def insert_site_data():
    print("> Inserting site data")
    for site_n, site_name in enumerate(site_data):
        print(f"> Filling out data for `{site_name}`")
        excel_template_layout = xw.Book(SOURCE_FILE)

        for sheet_name in excel_template_layout.sheet_names:
            print(f"> Filling out sheet `{sheet_name}`")
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
                sheet["C18"].value = sheet_data.get("ean_electricity", "")
                sheet["C19"].value = sheet_data.get("ean_gas", "")
                sheet["C20"].value = sheet_data.get("grid_operator_electricity", "")
                sheet["C21"].value = sheet_data.get("grid_operator_gas", "")

            if (
                stripped_sheet_name == strip_string("Koersvaste Ambitie")
                or stripped_sheet_name == strip_string("Eigen Vermogen")
                or stripped_sheet_name == strip_string("Gemeenschappelijke balans")
                or stripped_sheet_name == strip_string("Horizon aanvoer")
            ):
                for year in ["2021"] + list(sheet_data.keys()):
                    ROW = {
                        "2021": 4,
                        "2030": 10,
                        "2035": 13,
                        "2040": 16,
                        "2050": 19,
                    }

                    data = (
                        site_data[site_name]["base"]["2021"]
                        if year == "2021"
                        else sheet_data[year]
                    )

                    for i in range(0, 2):
                        dem_sup = "Demand" if i == 0 else "Supply"
                        electricity_anual = data[dem_sup].get("electricity_anual", "")
                        electricity_peak = data[dem_sup].get("electricity_peak", "")
                        electricity_generation_type = (
                            "WKK"
                            if electricity_anual > 0 and electricity_peak > 0
                            else ""
                        )
                        row = ROW[year] + i

                        sheet[f"K{row}"].value = data[dem_sup].get("co2_fossil", "")
                        sheet[f"L{row}"].value = data[dem_sup].get("co2_bio", "")
                        sheet[f"M{row}"].value = electricity_anual
                        sheet[f"N{row}"].value = electricity_peak
                        sheet[f"O{row}"].value = electricity_generation_type
                        sheet[f"P{row}"].value = data[dem_sup].get("natural_gas", "")
                        sheet[f"Q{row}"].value = data[dem_sup].get(
                            "hydrogen_more_than_98", ""
                        )
                        sheet[f"R{row}"].value = data[dem_sup].get(
                            "hydrogen_less_than_98", ""
                        )
                        sheet[f"S{row}"].value = data[dem_sup].get(
                            "heat_less_than_100", ""
                        )
                        sheet[f"T{row}"].value = data[dem_sup].get(
                            "heat_more_than_100", ""
                        )

            if stripped_sheet_name == strip_string("4. Flexibiliteit (begeleid)"):
                for year_number in site_data[site_name][stripped_sheet_name]:
                    for year_title in site_data[site_name][stripped_sheet_name][
                        year_number
                    ]:
                        ROW = {
                            "2021": {"base": 4},
                            "2030": {
                                "eigen_vermogen": 12,
                                "koersvaste_ambitie": 20,
                                "gemeenschappelijke_balans": 28,
                                "horizon_aanvoer": 36,
                            },
                            "2035": {
                                "eigen_vermogen": 44,
                                "koersvaste_ambitie": 52,
                                "gemeenschappelijke_balans": 60,
                                "horizon_aanvoer": 68,
                            },
                            "2040": {
                                "eigen_vermogen": 76,
                                "koersvaste_ambitie": 84,
                                "gemeenschappelijke_balans": 92,
                                "horizon_aanvoer": 100,
                            },
                            "2050": {
                                "eigen_vermogen": 108,
                                "koersvaste_ambitie": 116,
                                "gemeenschappelijke_balans": 124,
                                "horizon_aanvoer": 132,
                            },
                        }

                        for n, data in enumerate(
                            site_data[site_name][stripped_sheet_name][year_number][
                                year_title
                            ]
                        ):
                            row = ROW[year_number][year_title] + n
                            sheet[f"D{row}"].value = data.get("flex_means", "")
                            sheet[f"F{row}"].value = data.get("flexible_power", "")
                            sheet[f"E{row}"].value = data.get("availability", "")
                            sheet[f"H{row}"].value = data.get("operational_hours", "")
                            sheet[f"I{row}"].value = data.get("buffer_capacity", "")
                            sheet[f"J{row}"].value = data.get(
                                "performance_coefficient", ""
                            )
                            sheet[f"K{row}"].value = data.get("reduction_shift", "")
                            sheet[f"M{row}"].value = data.get("consecutive_hours", "")

        # Create output file
        os.makedirs(output_folder, exist_ok=True)
        filename = f"{output_folder}/{site_name}.xlsx"
        excel_template_layout.save(filename)
        print(f"Created `{filename}` ({site_n+1}/{len(site_data)})")
        excel_template_layout.close()


# Main
def main():
    time_start = time.time()

    # Extract data from the input file
    extract_site_data()

    # Put data in output files
    insert_site_data()

    time_end = time.time()
    duration = time_end - time_start
    print(
        f"Done, process took {math.floor(duration / 60)} minute(s) and {round(duration % 60)} seconds"
    )


if __name__ == "__main__":
    main()
