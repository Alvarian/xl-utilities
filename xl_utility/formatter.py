from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
wb = load_workbook('tests/demographics/main.xlsx')
ws = wb.active

import re, string


def separateNames(col_names):
    condition = range(1, ws.max_row)
    if not _find_column_by_name("lastname"):
        ws.insert_cols(0)
        ws["A1"].value = "First Name"

        ws.insert_cols(0)
        ws["A1"].value = "Last Name"

    has_first = str(_find_column_by_name("firstname"))
    has_last = str(_find_column_by_name("lastname"))
    
    def _separated_name(has_column_letter, list_of_test_payloads):
        split_cols = {
            "first": list(),
            "last": list()
        }
        for row_idx in condition:
            split_name = str(ws[has_column_letter + str(int(row_idx)+1)].value).split()
            split_cols["first"].append(split_name[0])
            
            if len(split_name) > 1:
                split_cols["last"].append(split_name[1])

        def _new_column(col_letter, col_position):
            altered_for_test = {}
            altered_for_test["column"] = ws[col_letter + "1"].value
            altered_for_test["data"] = list()

            for row_idx in condition:
                cell = split_cols[col_position][int(row_idx)-1].title()
                ws[col_letter + str(int(row_idx)+1)].value = cell
            
                altered_for_test["data"].append(cell)

            return altered_for_test

        list_of_test_payloads.append(_new_column(has_first, "first"))
        list_of_test_payloads.append(_new_column(has_last, "last"))

    return _parse_sheet_data(col_names, _separated_name)


def separateAddresses(col_names):
    def _separated_address(has_column_letter, list_of_test_payloads):
        def _alter_cell(row_idx):
            cell = str(ws[has_column_letter + str(int(row_idx)+1)].value)
            temp = re.compile("([0-9]+)([a-zA-Z]+)")

            if len(cell.split()) > 1:
                return cell
            else:
                altered_cell = " ".join(map(lambda x: x.capitalize(), temp.match(cell).groups()))
                ws[has_column_letter + str(int(row_idx)+1)].value = altered_cell

                return altered_cell

        altered_for_test = _alter_sheet_data(_alter_cell, has_column_letter)
        list_of_test_payloads.append(altered_for_test)

    return _parse_sheet_data(col_names, _separated_address)


def capitalizeFirstLetter(col_names):
    def _capitalized_first(has_column_letter, list_of_test_payloads):
        def _alter_cell(row_idx):
            cell = str(ws[has_column_letter + str(int(row_idx)+1)].value)
            altered_cell = cell.title()
            ws[has_column_letter + str(int(row_idx)+1)].value = altered_cell

            return altered_cell

        altered_for_test = _alter_sheet_data(_alter_cell, has_column_letter)
        list_of_test_payloads.append(altered_for_test)

    return _parse_sheet_data(col_names, _capitalized_first)


def capitalizeAll(col_names):
    def _capitalize_all(has_column_letter, list_of_test_payloads):
        def _alter_cell(row_idx):
            cell = str(ws[has_column_letter + str(int(row_idx)+1)].value)
            altered_cell = cell.upper()
            ws[has_column_letter + str(int(row_idx)+1)].value = altered_cell

            return altered_cell

        altered_for_test = _alter_sheet_data(_alter_cell, has_column_letter)
        list_of_test_payloads.append(altered_for_test)
    
    return _parse_sheet_data(col_names, _capitalize_all)

def _alter_sheet_data(_alter_cell, has_column_letter):
    altered_for_test = {}
    altered_for_test["column"] = ws[has_column_letter + "1"].value
    altered_for_test["data"] = list()

    for row_idx in range(1, ws.max_row):
        altered_cell = _alter_cell(row_idx)
        ws[has_column_letter + str(int(row_idx)+1)].value = altered_cell
        altered_for_test["data"].append(altered_cell)

    return altered_for_test

def _find_column_by_name(name=""):
    col_names = list(map(lambda x: x.replace(" ", "").lower(), list(map(lambda x: x.value, list(ws.iter_rows())[0]))))
    name_payload = re.sub(r"[^a-zA-Z0-9 ]", "", name.replace(" ", "").lower())
    
    return None if name_payload not in col_names else get_column_letter(col_names.index(name_payload)+1)


def _parse_sheet_data(col_names, handle_alterations):
    if type(col_names) is not list:
        raise TypeError("Arg with type {} is not list of strings".format(type(col_names)))

    list_of_test_payloads = list()

    # Just filters column names
    for name in col_names:
        has_column_letter = _find_column_by_name(name)
        if has_column_letter and ws[has_column_letter + "1"].value.replace(" ", "").lower() in list(map(lambda x: re.sub(r"[^a-zA-Z0-9 ]", "", x.replace(" ", "").lower()), col_names)):   
            handle_alterations(has_column_letter, list_of_test_payloads)

    if len(list_of_test_payloads) == 0:
        raise ValueError("Column names do not exist")

    return list_of_test_payloads

    