import sys, json
sys.path.insert(0, '')

from xl_utility.formatter import *



# print(capitalizeAll(["name", "Full names"]))
# print(separateNames(["name", "Full names"]))


def _assemble_comparison(col_names):
  mock_json = open("tests/demographics/main.json")
  mock_data = json.load(mock_json)
  print(mock_data)
  if len(col_names) > 1:
    payload = list()
    for col in col_names:
      if col in mock_data:
        for item in mock_data[col]:
          print(item)
          # payload.append({
            
          # })

    return payload
  
  return mock_data["name"]
# _assemble_comparison(["first name", "last name"])]

def test_separateNames():
  assert separateNames(["name", "Full names"]) == _assemble_comparison(["first name", "last name"])

def test_separateAddresses():
  assert separateAddresses(["address", "street address"]) == mock_data["street address"]

def test_capitalizeFirstLetter():
  assert capitalizeFirstLetter() == "capFirst"

def test_capitalizeAll():
  assert capitalizeAll(["name", "Full names"]) == mock_data["name"]