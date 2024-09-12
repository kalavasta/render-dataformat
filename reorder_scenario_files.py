# Rename scenario files to have a specific order in the scenario editor
# Works for the following folder scructure:
#
# scenarios/
# ├── 2030_EV/
# │   ├── 2030_CC_eigen_vermogen.json (becomes 1_2030_EV_CC.json)
# │   ├── 2030_eigen_vermogen.json (becomes 2_2030_EV_Dataformats.json)
# │   ├── 2030_EV_instellingen.json (becomes 3_2030_EV_Instellingen.json)

import os
import shutil

ROOT_DIR = "./scenarios"


def main():
    print(f"> Creating backup of `{ROOT_DIR}`")
    shutil.copytree(ROOT_DIR, f"{ROOT_DIR}_backup")
    print(f"> Backup created at `{ROOT_DIR}_backup`")

    for subdir, _, _ in os.walk(ROOT_DIR):
        if subdir == ROOT_DIR:
            continue

        subdir_name = subdir.replace(ROOT_DIR, "").replace("/", "")
        print(f"> Starting `{subdir_name}`")

        for _, _, files in os.walk(subdir):
            json_files = [file for file in files if file.endswith(".json")]
            json_files = sorted(json_files)

            for file in json_files:
                print(f"> Processing `{file}`")

                if (
                    subdir_name in file
                    and f"{subdir_name}_instellingen.json" not in file
                ):
                    print(f"> Already renamed, skipping...")
                    continue

                type = ""
                order = 0

                if "_CC_" in file:
                    type = "CC"
                    order = 1
                elif "_instellingen" in file:
                    type = "Instellingen"
                    order = 3
                else:
                    type = "Dataformats"
                    order = 2

                new_file_name = f"{order}_{subdir_name}_{type}.json"
                old_path = os.path.join(subdir, file)
                new_path = os.path.join(subdir, new_file_name)

                os.rename(old_path, new_path)
                print(f"> Renamed to `{new_file_name}`")


if __name__ == "__main__":
    main()
    print(f"Done!")
