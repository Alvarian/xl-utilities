import sys, json
from io import BytesIO, open
sys.path.insert(0, '')

from xl_utility.determinizer import *
import pytest

# EXCEL_FILE = open('tests/demographics/main.csv', 'rb')
EXCEL_FILE = open('tests/demographics/main.xlsx', 'rb')


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


# @pytest.mark.gg
# def test_guessGender():
#     assert guess_gender() == "guess"

# @pytest.mark.gu
# def test_generateUUID():
#     assert generate_uuid() == "genID"

@pytest.mark.im
def test_generateMockData():
    assert insert_mock_data(["Start Time"], "7:00am", EXCEL_FILE)["test_list"] == _assemble_comparison(["Start Time"], "insert mock")

    with pytest.raises(Exception) as e_info:
        insert_mock_data(42, "7:00am", EXCEL_FILE)

    with pytest.raises(Exception) as e_info:
        insert_mock_data(["name", 0], "7:00am", EXCEL_FILE)

    with pytest.raises(Exception) as e_info:
        insert_mock_data(["name", 0], ["7:00am", "9:30am"], EXCEL_FILE)
