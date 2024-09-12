import os
from functions import create_file
import json
import requests
import copy
import sys

URL = "https://beta.carbontransitionmodel.com"
DEFAULT_SCENARIO = "Base year"
ROOT_DIR = "./scenarios"

sessions = {}
default_scenario = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_SCENARIO
req_json = {
    "ScenarioID": "base",
    "default_scenario": default_scenario,
    "inputs": {},
    "outputs": ["SessionID"],
}


def main():

    for subdir, _, _ in os.walk(ROOT_DIR):
        if subdir == ROOT_DIR:
            continue

        subdir_name = subdir.replace(ROOT_DIR, "").replace("/", "")
        print(f"> Starting `{subdir_name}`")

        for _, _, files in os.walk(subdir):
            json_files = [file for file in files if file.endswith(".json")]
            json_files = sorted(json_files)

            if len(json_files) == 0:
                exit(f"Error: No JSON files found in {subdir}")

            print(
                f"> Found {len(json_files)} JSON files in `{subdir}, creating session`"
            )

            session_name = subdir.replace(ROOT_DIR, "").replace("/", "")
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

                    if "error" in res_json:
                        exit(f"Error: {res_json['error']}")

                    if res.status_code != 200:
                        exit(f"Error: {res_json}")

                    if session_id == "":
                        if "SessionID" not in res_json:
                            exit(
                                f"Error: Failed to create a session for `{session_name}`"
                            )

                        session_id = res_json["SessionID"]
                        print(f"> Successfully created a session with id: {session_id}")

            sessions[session_name] = {
                "name": session_name,
                "SessionID": session_id,
                "files": json_files,
            }
            print(f"> Session successfully created for `{session_name}`")

    create_file("./sessions.json", sessions)


if __name__ == "__main__":
    main()
