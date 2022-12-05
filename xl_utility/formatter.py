from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

from functools import partial
import re, string


def separateNames(col_names, excel_file):
    def _separated_name(col_name, list_of_test_payloads, ws, condition):
        if not _find_column_by_name("lastname", ws):
            ws.insert_cols(0)
            ws["A1"].value = "First Name"

            ws.insert_cols(0)
            ws["A1"].value = "Last Name"

        has_first = str(_find_column_by_name("firstname", ws))
        has_last = str(_find_column_by_name("lastname", ws))
        
        def _core(has_initial):
            split_cols = {
                "first": list(),
                "last": list()
            }
            for row_idx in condition:
                position = has_initial + str(int(row_idx)+1)
                cell = ws[position]
                split_name = str(cell.value).split()
                
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
        
        _validate_column(_core, col_name, col_names, ws)

    return _parse_sheet_data(col_names, _separated_name, excel_file)


def separateAddresses(col_names, excel_file):
    def _separated_address(col_name, list_of_test_payloads, ws, condition):
        def _alter_cell(row_idx, has_initial):
            cell = str(ws[has_initial + str(int(row_idx)+1)].value)
            temp = re.compile("([0-9]+)([a-zA-Z]+)")

            if len(cell.split()) > 1:
                return cell
            else:
                altered_cell = " ".join(map(lambda x: x.capitalize(), temp.match(cell).groups()))
                ws[has_initial + str(int(row_idx)+1)].value = altered_cell

                return altered_cell

        altered_for_test = _alter_sheet_data(_alter_cell, col_name, col_names, ws)
        list_of_test_payloads.append(altered_for_test)

    return _parse_sheet_data(col_names, _separated_address, excel_file)


def capitalizeFirstLetter(col_names, excel_file):
    def _capitalized_first(col_name, list_of_test_payloads, ws, condition):
        def _alter_cell(row_idx, has_initial):
            cell = str(ws[has_initial + str(int(row_idx)+1)].value)
            altered_cell = cell.title()
            ws[has_initial + str(int(row_idx)+1)].value = altered_cell

            return altered_cell

        altered_for_test = _alter_sheet_data(_alter_cell, col_name, col_names, ws)
        list_of_test_payloads.append(altered_for_test)

    return _parse_sheet_data(col_names, _capitalized_first, excel_file)


def capitalizeAll(col_names, excel_file):
    def _capitalize_all(col_name, list_of_test_payloads, ws, condition):
        def _alter_cell(row_idx, has_initial):
            cell = str(ws[has_initial + str(int(row_idx)+1)].value)
            altered_cell = cell.upper()
            ws[has_initial + str(int(row_idx)+1)].value = altered_cell

            return altered_cell

        altered_for_test = _alter_sheet_data(_alter_cell, col_name, col_names, ws)
        list_of_test_payloads.append(altered_for_test)
    
    return _parse_sheet_data(col_names, _capitalize_all, excel_file)


def _validate_column(_core, col_name, col_names, ws):
    has_initial = _find_column_by_name(col_name, ws)

    if has_initial and ws[has_initial + "1"].value.replace(" ", "").lower() in list(map(lambda x: re.sub(r"[^a-zA-Z0-9 ]", "", x.replace(" ", "").lower()), col_names)):   
        return _core(has_initial)
    else:
        raise KeyError("Column name not found")

def _alter_sheet_data(_alter_cell, col_name, col_names, ws):
    def _core(has_column_letter):
        altered_for_test = {}
        altered_for_test["column"] = ws[has_column_letter + "1"].value
        altered_for_test["data"] = list()

        for row_idx in range(1, ws.max_row):
            altered_cell = _alter_cell(row_idx, has_column_letter)
            ws[has_column_letter + str(int(row_idx)+1)].value = altered_cell
            altered_for_test["data"].append(altered_cell)

        return altered_for_test

    return _validate_column(_core, col_name, col_names, ws)

def _find_column_by_name(name, ws):
    col_names = list(map(lambda x: x.replace(" ", "").lower(), list(map(lambda x: x.value, list(ws.iter_rows())[0]))))
    name_payload = re.sub(r"[^a-zA-Z0-9 ]", "", name.replace(" ", "").lower())
    
    return None if name_payload not in col_names else get_column_letter(col_names.index(name_payload)+1)


def _parse_sheet_data(col_names, handle_alterations, excel_file):
    if type(col_names) is not list or not all(list(map(lambda x: type(x) == str, col_names))):
        raise TypeError("Arg with type {} is not list of strings".format(type(col_names)))

    list_of_test_payloads = list()

    wb = load_workbook(filename=excel_file, data_only=True)
    ws = wb.active
    
    condition = range(1, ws.max_row)
    
    # Filters column names
    for name in col_names:
        try:
            handle_alterations(name, list_of_test_payloads, ws, condition)
        except Exception as error:
            print(error)
            continue

    if len(list_of_test_payloads) == 0:
        raise ValueError("Column names do not exist")

    return list_of_test_payloads

    