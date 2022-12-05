import sys, json
sys.path.insert(0, '')

from xl_utility.formatter import *
import pytest


def _assemble_comparison(col_names, alter_name):
  mock_json = open("tests/demographics/main.json")
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
    _unit(42)

  with pytest.raises(Exception) as e_info:
    _unit(["name", 0])

  with pytest.raises(Exception) as e_info:
    _unit(["Constantinople", "icantspell"])


@pytest.mark.sn
def test_separateNames():
  assert separateNames(["name", "Full names"]) == _assemble_comparison(["First Name", "Last Name"], "seperate name")
  _main_raise_checkers(separateNames)

@pytest.mark.sa
def test_separateAddresses():
  assert separateAddresses(["address", "street address"]) == _assemble_comparison(["street address"], "seperate address")
  _main_raise_checkers(separateAddresses)

@pytest.mark.cf
def test_capitalizeFirstLetter():
  assert capitalizeFirstLetter(["name", "Full names"]) == _assemble_comparison(["name"], "capitalizeFirst")
  _main_raise_checkers(capitalizeFirstLetter)

@pytest.mark.ca
def test_capitalizeAll():
  assert capitalizeAll(["name", "Full names"]) == _assemble_comparison(["name", "Full names"], "capitalize")
  _main_raise_checkers(capitalizeAll)
