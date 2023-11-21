# Imports
import json
import os
import pandas as pd
import requests
import sys

# Constants
SHEETS = {
    "company_details": "1. Uitleg en Bedrijfsgegevens",
    "emissions_and_energy": "2. Emissies en energiebalansen",
    "projects": "3. Projecten",
    "flex": "4. Flexibiliteit (begeleid)",
}

YEARS = [
    "2030",
    "2035",
    "2040_decentrale_initiatieven",
    "2040_nationaal_leiderschap",
    "2040_europese_integratie",
    "2040_internationale_handel",
    "2050_decentrale_initiatieven",
    "2050_nationaal_leiderschap",
    "2050_europese_integratie",
    "2050_internationale_handel",
]

API_URL = "http://ctm-api-beta.eba-pamspfvv.eu-central-1.elasticbeanstalk.com/"

# Global variables
excel_folder = sys.argv[1]
json_folder = sys.argv[2]
sheet_data = {key: {} for key in YEARS}
sheet_data.update({"data": {}})
new_sites = {}

# Obtain list of sectors, clusters, and sites
response = requests.get(f"{API_URL}/api/getClusterInfo")
cc_data = response.json()


# Functions
def strip_string(string):
    string = (
        string.strip()
        .replace("&&", "####")
        .replace("-", "_")
        .replace(" ", "_")
        .replace("?", "")
        .replace("!", "")
        .replace(".", "")
        .replace("&", "")
        .replace(",", "")
        .replace("(", "")
        .replace(")", "")
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
        .replace("####", "&&")
    )
    return string.replace("__", "_")


def represents_int(s):
    try:
        int(s)
    except ValueError:
        return False
    else:
        return True


def extract_excel_data(excel_file, cc_data, new_count):
    key_prefix = ""
    error = False
    new_site = False
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
                print(
                    f"Industry for excel {excel_file} ({industry}) does not exist, it will be emitted from the inputs. \n\nPlease pick an industry from one of the following:\n",
                    cc_data["sectors"],
                    "\n\n\n",
                )
                return True, new_count, False
            if cluster not in cc_data["clusters"]:
                print(
                    f"Cluster for excel {excel_file} ({cluster}) does not exist, it will be emitted from the inputs. \n\nPlease pick a cluster from one of the following:\n",
                    cc_data["clusters"],
                    "\n\n\n",
                )
                return True, new_count, False
            if is_new_site:
                if name in cc_data["sites"]:
                    print(
                        f"Name for excel {excel_file} ({name}) already exists, site will be ignored.\n\n If you want to add a new site, please pick a name that does not match one of the following:\n",
                        cc_data["sites"],
                        "\n\n\n",
                    )
                    return True, new_count, False

                key_prefix = f"ldsh&&##new_cc_site{new_count}##"
                sheet_data["data"].update(
                    {
                        f"{key_prefix}_industry": industry,
                        f"{key_prefix}_cluster": cluster,
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
                        f"Name for excel {excel_file} ({name}) does not exist, site will be ignored.\n\n If you want to edit an existing site, please pick from one of the following:\n",
                        cc_data["sites"],
                        "\n\n\n",
                    )
                    return True, new_count, False

                key_prefix = strip_string(f"ldsh&&{industry}&&{cluster}&&{name}")

            sheet_data["data"][f"{key_prefix}&&ldsh_enabled"] = 1

            for row_n in range(7, 20):
                key = strip_string(
                    f"{key_prefix}&&{sheet_key}_{excel_content.iloc[row_n, 1]}"
                )
                sheet_data["data"].update({key: excel_content.iloc[row_n, 2]})

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

        elif sheet_key == "flex":
            year = ""
            year_suffix = ""
            current_flexibility = 1
            n_rows = excel_content[excel_content.columns[0]].count()
            for row_n in range(2, n_rows):
                if excel_content.iloc[row_n, 1] != "":
                    year = str(int(excel_content.iloc[row_n, 1]))
                    current_flexibility = 1
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
                current_flexibility += 1
    return error, new_count, new_site


def create_json_files():
    for year_key in YEARS:
        filename = f"{json_folder}/{str(year_key)}.json"
        os.makedirs(os.path.dirname(filename), exist_ok=True)

        with open(filename, "w") as f:
            json.dump(
                {**sheet_data["data"], **sheet_data[year_key], "new_sites": new_sites},
                f,
                indent=4,
            )
        print(f"Created {filename}")


# Main
def main():
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


if __name__ == "__main__":
    main()
