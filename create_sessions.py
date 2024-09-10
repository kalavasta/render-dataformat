import os
from functions import create_file

sessions = {}


def main():
    rootdir = "./scenarios"

    for subdir, _, _ in os.walk(rootdir):
        if subdir == rootdir:
            continue

        print(f"> Starting `{subdir}`")

        for _, _, files in os.walk(subdir):
            json_files = [file for file in files if file.endswith(".json")]

            if len(json_files) == 0:
                exit(f"Error: No JSON files found in {subdir}, skipping...")

            print(
                f"> Found {len(json_files)} JSON files in `{subdir}, creating session`"
            )

            session_name = subdir.replace(rootdir, "").replace("/", "")

            sessions[session_name] = {
                "name": session_name,
                "id": "",
                "files": json_files,
            }

            for file in json_files:
                print(f"> Processing `{file}`")

    print(sessions)
    create_file("./sessions.json", sessions)


if __name__ == "__main__":
    main()
