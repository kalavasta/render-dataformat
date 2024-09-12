# This script is for use with Cluster Analysis.xlsx: Site_resultaten_voor_upload

# Imports
import json
import os
import pandas as pd
import requests
import sys
import re
import glob
from functions import create_file, strip_string, represents_int

# Constants
SHEET = "Site_resultaten_voor_upload"  # Sheet name in the excel file

YEARS = [
    "2021",
    "2030",
    "2030_koersvaste_middenweg",
    "2030_eigen_toekomstbeeld_bedrijf",
    "2030_gezamenlijke_balans",
    "2030_horizon_aanvoer",
    "2030_eigen_vermogen",
    "2035",
    "2035_koersvaste_middenweg",
    "2035_eigen_toekomstbeeld_bedrijf",
    "2035_gezamenlijke_balans",
    "2035_horizon_aanvoer",
    "2035_eigen_vermogen",
    "2040_koersvaste_middenweg",
    "2040_eigen_toekomstbeeld_bedrijf",
    "2040_gezamenlijke_balans",
    "2040_horizon_aanvoer",
    "2040_eigen_vermogen",
    "2050_koersvaste_middenweg",
    "2050_eigen_toekomstbeeld_bedrijf",
    "2050_gezamenlijke_balans",
    "2050_horizon_aanvoer",
    "2050_eigen_vermogen",
]

API_URL = "https://carbontransitionmodel.com"

# Global variables
excel_folder = sys.argv[1]
json_folder = sys.argv[2]
sheet_data = {key: {} for key in YEARS}
sheet_data.update({"data": {}})
new_sites = {}
changes = []

# Obtain list of sectors, clusters, and sites
response = requests.get(f"{API_URL}/api/getClusterInfo/")
cc_data = response.json()

# Remove existing json files
json_files = glob.glob(os.path.join(json_folder, "*.json"))

for json_file in json_files:
    try:
        os.remove(json_file)
    except Exception as e:
        print(f"Error deleting {json_file}: {e}")


# Functions
def extract_excel_data(excel_file, cc_data, new_count):
    key_prefix = ""
    error = False
    new_site = {}
    included_new_sites = {}

    if SHEET not in pd.ExcelFile(excel_file).sheet_names:
        exit(f"Error: Sheet `{SHEET}` not found in `{excel_file}`")

    excel_content = pd.read_excel(excel_file, engine="openpyxl", sheet_name=SHEET)
    excel_content = excel_content.fillna("")
    n_rows = len(excel_content)

    for row_n in range(12, (n_rows)):
        if excel_content.iloc[row_n, 0] == "":
            print(f"> Row {row_n + 2}: Empty row, skipping...")
            continue

        industry = excel_content.iloc[row_n, 2]
        cluster = excel_content.iloc[row_n, 3]
        name = excel_content.iloc[row_n, 1]
        is_new_site = excel_content.iloc[row_n, 8].lower() == "nieuw"
        year = excel_content.iloc[row_n, 4]
        year_suffix = excel_content.iloc[row_n, 5]
        year_key = strip_string(f"{year} {year_suffix}")

        if industry not in cc_data["sectors"]:
            found = False

            for key in cc_data["sbi_codes"]:
                if str(industry) in cc_data["sbi_codes"][key]:
                    industry = key
                    found = True

            if not found:
                print(
                    f'Industry for excel {excel_file} (row {row_n + 2}: "{industry}") does not exist, it will be emitted from the inputs.\nPlease pick an industry from logs/industries.log or an SBI code from logs/sbi_codes.log\n'
                )
                return True, new_count, False

        if cluster not in cc_data["clusters"]:
            print(
                f'Cluster for excel {excel_file} (row {row_n + 2}: "{cluster}") does not exist, it will be emitted from the inputs.\nPlease pick a cluster from logs/clusters.log\n'
            )

            return True, new_count, False

        if is_new_site or name not in cc_data["sites"]:
            if is_new_site and name in cc_data["sites"]:
                print(
                    f'Name for excel {excel_file} (row {row_n + 2}: "{name}") already exists, site will be ignored.\nIf you want to add a new site, please pick a name that does not match one from logs/sites.log\n'
                )

                return True, new_count, False

            key_prefix = f"ldsh&&##new_cc_site{new_count}##"

            if name not in included_new_sites:
                to_change = f"##new_cc_site{new_count}##"
                included_new_sites[name] = key_prefix

                new_count += 1
                new_site[key_prefix] = {
                    "site": name,
                    "sector": industry,
                    "cluster": cluster,
                }

            else:
                key_prefix = included_new_sites[name]

            sheet_data[year_key].update(
                {
                    f"{key_prefix}&&industry": industry,
                    f"{key_prefix}&&cluster": cluster,
                    f"{key_prefix}&&company_details_locatie": name,
                }
            )

        else:
            if name not in cc_data["sites"]:
                print(
                    f'Name for excel {excel_file} (row {row_n + 2}: "{name}") does not exist, but site will be added.\nIf you want to edit an existing site, please pick from logs/sites.log\n',
                )

            key_prefix = strip_string(f"ldsh&&{industry}&&{cluster}&&{name}")
            to_change = strip_string(f"{industry}&&{cluster}&&{name}")

        changes.append(to_change)

        sheet_data[year_key][f"{key_prefix}&&ldsh_enabled"] = 1
        n_cols = len(excel_content.columns)

        for col_n in range(8, n_cols):
            col_name = excel_content.iloc[7, col_n]
            if col_name == "":
                continue
            key = strip_string(f"{key_prefix}&&{col_name}")
            value = excel_content.iloc[row_n, col_n]

            sheet_data[year_key].update({key: value})

    if strip_string(name) in cc_data["cc_sites"]:
        changes.append(key_prefix.replace("ldsh&&", ""))

    return error, new_count, new_site


def create_json_files():
    for year_key in YEARS:
        filename = f"{json_folder}/{str(year_key)}.json"
        year_data = sheet_data[year_key]

        create_file(
            filename,
            {
                **year_data,
                "new_sites": new_sites,
                "changes": list(set(changes)),
            },
        )


# Main
def main():
    new_count = 50

    for file in os.listdir(excel_folder):
        if file.endswith(".xlsx") and file[0] != "~":
            excel_file = f"{excel_folder}/{file}"
            [error, new_count, new_site] = extract_excel_data(
                excel_file, cc_data, new_count
            )

            if error:
                exit()

            if new_site:
                new_sites.update(new_site)

    create_json_files()
    create_file("logs/industries.log", cc_data["sectors"])
    create_file("logs/clusters.log", cc_data["clusters"])
    create_file(
        "logs/sites.log",
        [
            item
            for item in cc_data["sites"]
            if not re.match(r"##new_cc_site\d+##", item)
        ],
    )
    create_file("logs/sbi_codes.log", cc_data["sbi_codes"])


if __name__ == "__main__":
    main()
