import sys, json
sys.path.insert(0, '')

from xl_utility.formatter import *


def _create_altered_data(method_name, payload_of_altered_data):
    for altered_data in payload_of_altered_data:
        print(altered_data)
        with open("tests/demographics/main.json",'r+') as file:
            # First we load existing data into a dict.
            file_data = json.load(file)

            # Join new_data with file_data inside emp_details
            if not file_data.get(altered_data["column"]):
                file_data[altered_data["column"]] = {}

            file_data[altered_data["column"]][method_name] = altered_data["data"]

            # Sets file's current position at offset.
            file.seek(0)
            
            # convert back to json.
            json.dump(file_data, file, indent = 4)

_create_altered_data("seperate name", separateNames(["name", "Full names"]))
_create_altered_data("capitalize", capitalizeAll(["name", "Full names", "first name", "last name"]))
_create_altered_data("capitalizeFirst", capitalizeFirstLetter(["name", "Full names"]))
_create_altered_data("seperate address", separateAddresses(["address", "street address"]))

