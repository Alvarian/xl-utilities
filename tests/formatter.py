import sys, json
sys.path.insert(0, '')

from xl_utility.formatter import *
import pytest

EXCEL_FILE = 'tests/demographics/main.xlsx'


def _assemble_comparison(col_names, alter_name):
  mock_json = open('tests/demographics/main.json')
  mock_data = json.load(mock_json)
  
  payload = list()
  for col in col_names:
    if col in mock_data:
      payload.append({
        "column": col,
        "data": mock_data[col][alter_name]
      })

  return payload
  
def _main_raise_checkers(_unit):
  with pytest.raises(Exception) as e_info:
    _unit(42, EXCEL_FILE)

  with pytest.raises(Exception) as e_info:
    _unit(["name", 0], EXCEL_FILE)
    

@pytest.mark.sn
def test_separateNames():
  assert separate_names(["name", "Full names"], EXCEL_FILE)["test_list"] == _assemble_comparison(["First Name", "Last Name"], "seperate name")
  _main_raise_checkers(separate_names)

  assert separate_names(["street address"], EXCEL_FILE)["exception"] == "Non text cells are forbidden in this function. -street address is rejected! "
  assert separate_names(["email address"], EXCEL_FILE)["exception"] == "Email cells are forbidden in this function. -email address is rejected! "
  assert separate_names(["phone number"], EXCEL_FILE)["exception"] == "Non text cells are forbidden in this function. -phone number is rejected! "


@pytest.mark.sa
def test_separateAddresses():
  assert separate_addresses(["address", "street address"], EXCEL_FILE)["test_list"] == _assemble_comparison(["street address"], "seperate address")
  _main_raise_checkers(separate_addresses)

  assert separate_addresses(["email address"], EXCEL_FILE)["exception"] == "Email cells are forbidden in this function. -email address is rejected! "
  assert separate_addresses(["phone number"], EXCEL_FILE)["exception"] == "Number cells are forbidden in this function. -phone number is rejected! "


@pytest.mark.cf
def test_capitalizeFirstLetter():
  assert capitalize_firstLetter(["name", "Full names"], EXCEL_FILE)["test_list"] == _assemble_comparison(["name"], "capitalizeFirst")
  _main_raise_checkers(capitalize_firstLetter)

  assert capitalize_firstLetter(["phone number"], EXCEL_FILE)["exception"] == "Number cells are forbidden in this function. -phone number is rejected! "


@pytest.mark.ca
def test_capitalizeAll():
  assert capitalize_all(["name", "Full names"], EXCEL_FILE)["test_list"] == _assemble_comparison(["name", "Full names"], "capitalize")
  _main_raise_checkers(capitalize_all)

  assert capitalize_all(["phone number"], EXCEL_FILE)["exception"] == "Number cells are forbidden in this function. -phone number is rejected! "
  