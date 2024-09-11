import os
from functions import create_file
import json
import requests
import copy
import sys

URL = "https://beta.carbontransitionmodel.com"
DEFAULT_SCENARIO = "Base year"

sessions = {}
default_scenario = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SCENARIO
req_json = {
    "ScenarioID": "base",
    "default_scenario": default_scenario,
    "inputs": {},
    "outputs": ["SessionID"],
}


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
            session_id = ""

            for file in json_files:
                print(f"> Processing `{file}`")
                filepath = os.path.join(subdir, file)

                with open(filepath) as f:
                    if session_id == "":
                        print(f"> Creating session for `{session_name}`")
                    else:
                        print(f"> Updating session for `{session_name}`")
                    new_inputs = json.load(f)

                    req = copy.deepcopy(req_json)
                    req["inputs"] = new_inputs
                    if session_id != "":
                        req["SessionID"] = session_id
                    res = requests.post(url=f"{URL}/api/", json=req)
                    res_json = res.json()

                    if session_id == "":
                        print(
                            f"> Successfully created a session with id: {res_json['SessionID']}"
                        )

                    session_id = res_json["SessionID"]

            sessions[session_name] = {
                "name": session_name,
                "id": session_id,
                "files": json_files,
            }
            print(f"> Session successfully created for `{session_name}`")

    create_file("./sessions.json", sessions)


if __name__ == "__main__":
    main()
