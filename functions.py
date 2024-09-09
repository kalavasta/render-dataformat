import os
import json


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


def create_file(filename, data):
    os.makedirs(os.path.dirname(filename), exist_ok=True)
    with open(filename, "w") as f:
        json.dump(
            data,
            f,
            indent=4,
        )
    print(f"Created {filename}")
