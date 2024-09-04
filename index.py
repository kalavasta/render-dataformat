# Imports
import json
import os
import pandas as pd
import requests
import sys
import re
import openpyxl
import glob

# Constants
SHEETS = {
    "company_details": "1. Uitleg en Bedrijfsgegevens",
    "emissions_and_energy": "2. Emissies en energiebalansen",
    "projects": "3. Projecten",
    "flex": "4. Flexibiliteit (begeleid)",
}

SHEETS_NEW_SETUP = {
    "company_details": "1. Uitleg en Bedrijfsgegevens",
    "koersvaste_ambitie": "Koersvaste Ambitie",
    "eigen_vermogen": "Eigen Vermogen",
    "gemeenschappelijke_balans": "Gemeenschappelijke balans",
    "horizon_aanvoer": "Horizon aanvoer",
    "flex": "4. Flexibiliteit (begeleid)"
}

YEARS = [
    "2021",
    "2030",
    "2035",
    "2040_eigen_toekomstbeeld_bedrijf",
    "2050_eigen_toekomstbeeld_bedrijf",
    "2040_decentrale_initiatieven",
    "2040_nationaal_leiderschap",
    "2040_europese_integratie",
    "2040_internationale_handel",
    "2050_decentrale_initiatieven",
    "2050_nationaal_leiderschap",
    "2050_europese_integratie",
    "2050_internationale_handel",
]

years_option = ["2030", "2035", "2040", "2050"]

YEARS_NEW = [f"{year}_{new_sheet}" for year in years_option for new_sheet in SHEETS_NEW_SETUP if new_sheet != "flex" and new_sheet != "company_details"] + ["2021"]

API_URL = "https://carbontransitionmodel.com"

# Global variables
excel_folder = sys.argv[1]
json_folder = sys.argv[2]
sheet_data = {key: {} for key in YEARS + YEARS_NEW}
sheet_updated = {key: key=="2021" for key in YEARS + YEARS_NEW}
sheet_data.update({"data": {}})
new_sites = {}
changes = []

# Obtain list of sectors, clusters, and sites
response = requests.get(f"{API_URL}/api/getClusterInfo/")
cc_data = response.json()


# Functions
def strip_string(string):
    string = (
        string.strip()
        .replace("&&", "#replace#")
        .replace("-", "_")
        .replace(" ", "_")
        .replace("?", "")
        .replace("!", "")
        .replace(".", "")
        .replace("&", "")
        .replace(",", "")
        .replace("(", "")
        .replace(")", "")
        .replace("'", "")
        .replace("<", "_less_than_")
        .replace(">", "_more_than_")
        .replace("%", "")
        .replace(":", "")
        .replace("€", "")
        .replace("ë", "e")
        .replace("ö", "o")
        .replace("/", "_")
        .replace("\n", "_")
        .replace("__", "_")
        .rstrip("_")
        .lower()
        .replace("#replace#", "&&")
    )
    return string.replace("__", "_")


def represents_int(s):
    try:
        int(s)
    except ValueError:
        return False
    else:
        return True


def create_file(filename, data):
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    with open(filename, "w") as f:
        json.dump(
            data,
            f,
            indent=4,
        )
    print(f"Created {filename}")


def extract_excel_data(excel_file, cc_data, new_count):
    key_prefix = ""
    error = False
    new_site = False

    wb = openpyxl.load_workbook(excel_file)

    sheets = wb.sheetnames
    if not (set(sheets) - {"1. Uitleg en Bedrijfsgegevens","4. Flexibiliteit (begeleid)"} & set(SHEETS_NEW_SETUP.values())):
        for sheet_key, sheet_value in SHEETS.items():
            excel_content = pd.read_excel(
                excel_file, engine="openpyxl", sheet_name=sheet_value
            )
            excel_content = excel_content.fillna("")

            if sheet_key == "company_details":
                industry = excel_content.iloc[14, 2]
                cluster = excel_content.iloc[15, 2]
                name = excel_content.iloc[13, 2]
                is_new_site = excel_content.iloc[10, 2].lower() == "nieuw"

                if industry not in cc_data["sectors"]:
                    found = False
                    for key in cc_data["sbi_codes"]:
                        if str(industry) in cc_data["sbi_codes"][key]:
                            industry = key
                            found = True

                    if not found:
                        print(
                            f"Industry for excel {excel_file} ({industry}) does not exist, it will be emitted from the inputs.\nPlease pick an industry from logs/industries.log or an SBI code from logs/sbi_codes.log\n"
                        )
                        return True, new_count, False

                if cluster not in cc_data["clusters"]:
                    print(
                        f"Cluster for excel {excel_file} ({cluster}) does not exist, it will be emitted from the inputs.\nPlease pick a cluster from logs/clusters.log\n"
                    )

                    return True, new_count, False

                if is_new_site or name not in cc_data["sites"]:
                    if is_new_site and name in cc_data["sites"]:
                        print(
                            f"Name for excel {excel_file} ({name}) already exists, site will be ignored.\nIf you want to add a new site, please pick a name that does not match one from logs/sites.log\n"
                        )

                        return True, new_count, False

                    key_prefix = f"ldsh&&##new_cc_site{new_count}##"
                    sheet_data["data"].update(
                        {
                            f"{key_prefix}&&industry": industry,
                            f"{key_prefix}&&cluster": cluster,
                        }
                    )
                    new_count += 1
                    new_site = {
                        key_prefix: {
                            "site": name,
                            "sector": industry,
                            "cluster": cluster,
                        }
                    }
                else:
                    if name not in cc_data["sites"]:
                        print(
                            f"Name for excel {excel_file} ({name}) does not exist, but site will be added.\nIf you want to edit an existing site, please pick from logs/sites.log\n",
                        )

                    key_prefix = strip_string(f"ldsh&&{industry}&&{cluster}&&{name}")

                sheet_data["data"][f"{key_prefix}&&ldsh_enabled"] = 1

                for row_n in range(7, 20):
                    key = strip_string(
                        f"{key_prefix}&&{sheet_key}_{excel_content.iloc[row_n, 1]}"
                    )
                    value = excel_content.iloc[row_n, 2]
                    sheet_data["data"].update({key: value})

                if strip_string(name) in cc_data["cc_sites"]:
                    changes.append(key_prefix.replace("ldsh&&", ""))

            elif sheet_key == "emissions_and_energy":
                year = ""
                year_suffix = ""
                for row_n in range(3, 46):
                    if excel_content.iloc[row_n, 1] != "":
                        year = (
                            str(int(excel_content.iloc[row_n, 1]))
                            if represents_int(excel_content.iloc[row_n, 1])
                            else ""
                        )
                    if year == "":
                        continue
                    if excel_content.iloc[row_n, 2]:
                        year_suffix = str(excel_content.iloc[row_n, 2])
                    row_preheader = "base" if year == "2021" else "future"
                    row_header = (
                        "production"
                        if "Productie" in excel_content.iloc[row_n, 4]
                        else "demand"
                    )
                    if row_header == "" or excel_content.iloc[row_n, 4] == "":
                        continue
                    col_preheader = ""
                    for col_n in range(5, 34):
                        if excel_content.iloc[1, col_n] != "":
                            col_preheader = excel_content.iloc[1, col_n]
                        col_header = excel_content.iloc[2, col_n]
                        key = strip_string(
                            f"{key_prefix}&&{sheet_key}_{col_preheader}_{col_header}_{row_header}_{row_preheader}"
                        )

                        if excel_content.iloc[row_n, col_n] != "":
                            year_key = strip_string(f"{year}_{year_suffix}")
                            data_key = "data" if year == "2021" else year_key
                            sheet_data[data_key].update(
                                {key: excel_content.iloc[row_n, col_n]}
                            )
                            sheet_updated[data_key] = True


            elif sheet_key == "flex":
                year = ""
                year_suffix = ""
                current_flexibility = 1
                n_rows = excel_content[excel_content.columns[0]].count()
                for row_n in range(2, n_rows):
                    if (
                        excel_content.iloc[row_n, 1] != ""
                        or excel_content.iloc[row_n, 2] != ""
                    ):
                        current_flexibility = 1
                    if excel_content.iloc[row_n, 1] != "":
                        year = str(int(excel_content.iloc[row_n, 1]))
                    if excel_content.iloc[row_n, 2]:
                        year_suffix = str(excel_content.iloc[row_n, 2])
                    row_preheader = "base" if year == "2021" else "future"
                    row_header = f"flexibility_{str(current_flexibility)}"
                    col_preheader = ""
                    for col_n in range(3, 14):
                        if excel_content.iloc[0, col_n]:
                            col_preheader = excel_content.iloc[0, col_n]
                        col_header = excel_content.iloc[1, col_n]
                        key = strip_string(
                            f"{key_prefix}&&{sheet_key}_{col_preheader}_{col_header}_{row_header}_{row_preheader}"
                        )

                        if excel_content.iloc[row_n, col_n] != "":
                            year_key = strip_string(f"{year} {year_suffix}")
                            data_key = "data" if year == "2021" else year_key
                            sheet_data[data_key].update(
                                {key: excel_content.iloc[row_n, col_n]}
                            )
                            sheet_updated[data_key] = True
                    current_flexibility += 1
    else:
        for sheet_key, sheet_value in SHEETS_NEW_SETUP.items():
            excel_content = pd.read_excel(
                excel_file, engine="openpyxl", sheet_name=sheet_value
            )
            excel_content = excel_content.fillna("")
            if sheet_key == "company_details":
                industry = excel_content.iloc[14, 2]
                cluster = excel_content.iloc[15, 2]
                name = excel_content.iloc[13, 2]
                is_new_site = excel_content.iloc[10, 2].lower() == "nieuw"

                if industry not in cc_data["sectors"]:
                    found = False
                    for key in cc_data["sbi_codes"]:
                        if str(industry) in cc_data["sbi_codes"][key]:
                            industry = key
                            found = True

                    if not found:
                        print(
                            f"Industry for excel {excel_file} ({industry}) does not exist, it will be emitted from the inputs.\nPlease pick an industry from logs/industries.log or an SBI code from logs/sbi_codes.log\n"
                        )
                        return True, new_count, False

                if cluster not in cc_data["clusters"]:
                    print(
                        f"Cluster for excel {excel_file} ({cluster}) does not exist, it will be emitted from the inputs.\nPlease pick a cluster from logs/clusters.log\n"
                    )

                    return True, new_count, False

                if is_new_site or name not in cc_data["sites"]:
                    if is_new_site and name in cc_data["sites"]:
                        print(
                            f"Name for excel {excel_file} ({name}) already exists, site will be ignored.\nIf you want to add a new site, please pick a name that does not match one from logs/sites.log\n"
                        )

                        return True, new_count, False

                    key_prefix = f"ldsh&&##new_cc_site{new_count}##"
                    sheet_data["data"].update(
                        {
                            f"{key_prefix}&&industry": industry,
                            f"{key_prefix}&&cluster": cluster,
                        }
                    )
                    new_count += 1
                    new_site = {
                        key_prefix: {
                            "site": name,
                            "sector": industry,
                            "cluster": cluster,
                        }
                    }
                else:
                    if name not in cc_data["sites"]:
                        print(
                            f"Name for excel {excel_file} ({name}) does not exist, but site will be added.\nIf you want to edit an existing site, please pick from logs/sites.log\n",
                        )

                    key_prefix = strip_string(f"ldsh&&{industry}&&{cluster}&&{name}")

                sheet_data["data"][f"{key_prefix}&&ldsh_enabled"] = 1

                for row_n in range(7, 20):
                    key = strip_string(
                        f"{key_prefix}&&{sheet_key}_{excel_content.iloc[row_n, 1]}"
                    )
                    value = excel_content.iloc[row_n, 2]
                    sheet_data["data"].update({key: value})

                if strip_string(name) in cc_data["cc_sites"]:
                    changes.append(key_prefix.replace("ldsh&&", ""))
            elif sheet_key == "flex":
                year = ""
                year_suffix = ""
                current_flexibility = 1

                n_rows = excel_content[excel_content.columns[0]].count()
                for row_n in range(2, len(excel_content)):
                    if (
                        excel_content.iloc[row_n, 1] != ""
                        or excel_content.iloc[row_n, 2] != ""
                    ):
                        current_flexibility = 1
                    if excel_content.iloc[row_n, 1] != "":
                        year = str(int(excel_content.iloc[row_n, 1]))

                    if excel_content.iloc[row_n, 2]:
                        year_suffix = str(excel_content.iloc[row_n, 2])

                    row_preheader = "base" if year == "2021" else "future"
                    row_header = f"flexibility_{str(current_flexibility)}"
                    col_preheader = ""
                    for col_n in range(3, 14):
                        if excel_content.iloc[0, col_n]:
                            col_preheader = excel_content.iloc[0, col_n]
                        col_header = excel_content.iloc[1, col_n]
                        key = strip_string(
                            f"{key_prefix}&&{sheet_key}_{col_preheader}_{col_header}_{row_header}_{row_preheader}"
                        )

                        if excel_content.iloc[row_n, col_n] != "":
                            year_key = strip_string(f"{year} {year_suffix}")
                            data_key = "data" if year == "2021" else year_key
                            sheet_data[data_key].update(
                                {key: excel_content.iloc[row_n, col_n]}
                            )
                            sheet_updated[data_key] = True
                    current_flexibility += 1
            else:
                year = ""
                year_suffix = sheet_key
                check_ending = False
                row_n = 2
                while not check_ending or check_ending != "Verhaallijnen worden begeleid uitgevraagd.":
                    if excel_content.iloc[row_n, 1] != "":
                        year = (
                            str(int(excel_content.iloc[row_n, 1]))
                            if represents_int(excel_content.iloc[row_n, 1])
                            else ""
                        )
                    if year == "":
                        check_ending = excel_content.iloc[row_n, 1]
                        row_n += 1
                        continue
                    
                    row_preheader = "base" if year == "2021" else "future"
                    row_header = (
                        "production"
                        if "Productie" in excel_content.iloc[row_n, 4]
                        else "demand"
                    )
                    if row_header == "" or excel_content.iloc[row_n, 4] == "":
                        check_ending = excel_content.iloc[row_n, 1]
                        row_n += 1
                        continue

                    col_preheader = ""
                    col_n = 5
                    
                    while excel_content.shape[1] < col_n and excel_content.iloc[1, col_n] != "" :
                        if excel_content.iloc[0, col_n] != "":
                            col_preheader = excel_content.iloc[0, col_n]

                        col_header = excel_content.iloc[1, col_n]

                        key = strip_string(
                            f"{key_prefix}&&emissions_and_energy_{col_preheader}_{col_header}_{row_header}_{row_preheader}"
                        )

                        if excel_content.iloc[row_n, col_n] != "":
                            year_key = strip_string(f"{year}_{year_suffix}")
                            data_key = 'data' if year == "2021" else year_key

                            sheet_data[data_key].update(
                                {key: excel_content.iloc[row_n, col_n]}
                            )
                            sheet_updated[data_key] = True
                        col_n += 1
                    row_n += 1

                    check_ending = excel_content.iloc[row_n, 1]


    return error, new_count, new_site


def create_json_files():
    for year_key in YEARS + YEARS_NEW:
        if sheet_updated[year_key]:
            filename = f"{json_folder}/{str(year_key)}.json"
            year_data = sheet_data[year_key]
            if year_key == "2021":
                year_data = {}
                for key, value in sheet_data["data"].items():
                    new_key = re.sub("_base$", "_future", key)
                    year_data[new_key] = value

            create_file(
                filename,
                {
                    **sheet_data["data"],
                    **year_data,
                    "new_sites": new_sites,
                    "changes": changes,
                },
            )

def remove_contents(folder_path, extension):
    files = glob.glob(os.path.join(folder_path, f'*{extension}'))
    for file_path in files:
        try:
            os.remove(file_path)
        except Exception as e:
            print(f"Error removing {file_path}: {e}")


# Main
def main():
    remove_contents("json", ".json")
    new_count = 1
    for file in os.listdir(excel_folder):
        if file.endswith(".xlsx") and file[0] != "~":
            excel_file = f"{excel_folder}/{file}"
            [error, new_count, new_site] = extract_excel_data(
                excel_file, cc_data, new_count
            )
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
