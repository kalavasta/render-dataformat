# Render dataformat

This script will convert the excel files to json files that can be used in the transition model website.

The script will create a json file in the json folder for each year in the excel sheets. The json file will contain all the data from the excel file for that year, as well as all the data from the excel file for the base year (2021). The exported json files can be used in transition model website by uploading one of them via "Import Dataformats".

## Installation

Clone the repository to your local machine using the following command:

```
git clone https://github.com/kalavasta/render-dataformat.git
```

Install the required packages:

```
pip install -r requirements.txt
```

## Usage

Run the following script. Replace `<excel_folder>` with the folder containing the excel files and `<json_folder>` with the folder where you want to save the json files. The script will create the json folder if it does not exist.

```
python index.py <excel_folder> <json_folder>
```

This example script will use the excel and json folder at the root of this project.

```
python index.py excel json
```

## Implementation

When the JSON files have been generated, they can be imported into a session or scenario. This is illustrated in the `import_example.py` file, located in the root folder of this repository.
This works in the following manner. The first step is to create a session from the scenario that you want to use. The default is the `base` scenario, which contains the default values of a scenario. Next, the value of the created JSON file in the corresponding session, and outputs can be requested. Afterwards, the session is deleted, this step is optional.

## License

See the [LICENSE](LICENSE) file for license rights and limitations (MIT).
