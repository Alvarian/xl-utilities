from xl_utility import formatter, determinizer
from xl_utility.formatter import *
from xl_utility.determinizer import *

from inspect import getmembers, isfunction
import json

from io import BytesIO
from flask import Flask, request, render_template, send_file, session
import names

app = Flask(__name__)
app.debug = True
app.secret_key = "development key"


@app.route("/")
def index():
    omitted_functions = ["NamedTemporaryFile", "load_workbook", "get_column_letter"]

    formatter_features = list(filter(lambda x: bool(x[0].split("_")[0]) and not x[0].split()[0] in omitted_functions, getmembers(formatter, isfunction)))
    determinizer_features = list(filter(lambda x: bool(x[0].split("_")[0]) and not x[0].split()[0] in omitted_functions, getmembers(determinizer, isfunction)))

    feature_names = formatter_features + determinizer_features
    return render_template("index.html", features=feature_names)

@app.route("/events", methods = ["POST"])
def get_event_preset():
    EXCEL_FILE = BytesIO(request.files["file"].read())

    preset_payload = json.loads(request.form["preset_fields"])
    formatted_sn = separate_names(preset_payload["columns"][0], EXCEL_FILE)
    formatted_start_time = insert_mock_data(["Start Time"], preset_payload["Start Time"], BytesIO(formatted_sn["buffer"]))
    formatted_end_time = insert_mock_data(["End Time"], preset_payload["End Time"], BytesIO(formatted_start_time["buffer"]))

    final_exception = formatted_sn["exception"] + formatted_start_time["exception"] + formatted_end_time["exception"] if not all([formatted_sn["exception"], formatted_start_time["exception"], formatted_end_time["exception"]]) else None

    return send_file(
        BytesIO(formatted_end_time["buffer"]), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), final_exception

@app.route("/insert_mock_data", methods = ["POST"])
def get_mock_column():
    def _generate_ran_names():
        mock_list = list()
        for idx in range(0, int(request.form["max_row"])+1):
            mock_list.append(names.get_full_name())

        return {
            "column": ["random names"],
            "data": mock_list
        }

    def _assemble_custom_clone():
        inputs = request.form["custom_inputs"].split(",")
        return {
            "column": [inputs[0]],
            "data": inputs[1]
        }

    switcher = {
        "get_random_names": _generate_ran_names,
        "custom": _assemble_custom_clone
    }
    EXCEL_FILE = BytesIO(request.files["file"].read())
    module_args = switcher.get(request.form["mock_selected"], None)()

    has_altered_sheet = insert_mock_data(module_args["column"], module_args["data"], EXCEL_FILE)

    return send_file(
        BytesIO(has_altered_sheet["buffer"]), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), has_altered_sheet["exception"] or None

@app.route("/capitalize_all", methods = ["POST"])
def get_capitalized_column():
    EXCEL_FILE = BytesIO(request.files["file"].read())

    has_altered_sheet = capitalize_all(request.form["columns_selected"].split(","), EXCEL_FILE)
    return send_file(
        BytesIO(has_altered_sheet["buffer"]), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), has_altered_sheet["exception"] or None

@app.route("/capitalize_firstletter", methods = ["POST"])
def get_capitalized_first():
    EXCEL_FILE = BytesIO(request.files["file"].read())

    has_altered_sheet = capitalize_firstLetter(request.form["columns_selected"].split(","), EXCEL_FILE)
    return send_file(
        BytesIO(has_altered_sheet["buffer"]), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), has_altered_sheet["exception"] or None

@app.route("/separate_addresses", methods = ["POST"])
def get_separated_addresses():
    EXCEL_FILE = BytesIO(request.files["file"].read())

    has_altered_sheet = separate_addresses(request.form["columns_selected"].split(","), EXCEL_FILE)
    return send_file(
        BytesIO(has_altered_sheet["buffer"]), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), has_altered_sheet["exception"] or None

@app.route("/separate_names", methods = ["POST"])
def get_separated_names():
    EXCEL_FILE = BytesIO(request.files["file"].read())

    has_altered_sheet = separate_names(request.form["columns_selected"].split(","), EXCEL_FILE)
    return send_file(
        BytesIO(has_altered_sheet["buffer"]), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), has_altered_sheet["exception"] or None
