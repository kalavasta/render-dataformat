import requests
import json as js

url = "https://beta.carbontransitionmodel.com"

default_scenario = "Base year"

# First, request a session for the scenario you are trying to access
reqJson = {
    "ScenarioID": "base",
    "default_scenario": default_scenario,
    "inputs": {},
    "outputs": ["SessionID"],
}

resp = requests.post(url=f"{url}/api/", json=reqJson)

respJson = resp.json()
session_id = respJson["SessionID"]
print("Successfully created a session with id:", session_id)

# import the json file of the data format
with open("json/2035.json") as f:
    inputs = js.load(f)


# Send the new inputs and request outputs (for faster calculation times, include 'MSGraphSession', not necessary)
outputs = ["total_ctm_physical_indirect_emissions_dashboard"]
reqJson = {
    "SessionID": session_id,
    "default_scenario": default_scenario,
    "inputs": inputs,
    "outputs": outputs,
    "MSGraphSession": respJson["MSGraphSession"],
}
resp = requests.post(url=f"{url}/api/", json=reqJson)

outJson = resp.json()

# only show requested outputs
outputDic = {}
for item in outputs:
    outputDic[item] = outJson["output_values"][item]

print("The values for the requested outputs are:", outputDic)

# delete session after usage
reqJson = {"SessionID": session_id, "special": ["deleteSession"]}
resp = requests.post(url=f"{url}/api/", json=reqJson)
print(resp.text)
