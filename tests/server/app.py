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

class pipe:
    def __init__(self, value, func=None):
        self.value = value
        self.func = func
    def __getitem__(self, func):
        return pipe(self.value, func)        
    def __call__(self, *args, **kwargs):
        print(self.value)
        if type(self.value) == dict and self.value["exception"]:
            raise ValueError(self.value["exception"])

        return pipe(self.func(self.value, *args, **kwargs))
    def __repr__(self):
        return 'pipe(%s, %s)' % (self.value, self.func)



@app.route("/")
def index():
    omitted_functions = ["NamedTemporaryFile", "load_workbook", "get_column_letter"]

    formatter_features = list(filter(lambda x: bool(x[0].split("_")[0]) and not x[0].split()[0] in omitted_functions, getmembers(formatter, isfunction)))
    determinizer_features = list(filter(lambda x: bool(x[0].split("_")[0]) and not x[0].split()[0] in omitted_functions, getmembers(determinizer, isfunction)))

    feature_names = formatter_features + determinizer_features
    return render_template("index.html", features=feature_names)

@app.route("/events", methods = ["POST"])
def get_event_preset():
    EXCEL_FILE = {'buffer': request.files["file"], 'exception': ''}
    preset_payload = json.loads(request.form["preset_fields"])

    try:
        event_formatted = (pipe(EXCEL_FILE)
            [separate_names](preset_payload["columns"][0])
            [insert_mock_data](["Start Time"], preset_payload["Start Time"])
            [insert_mock_data](["End Time"], preset_payload["End Time"])
        .value)

        return send_file(
            BytesIO(event_formatted["buffer"].read()), 
            download_name="new_sheet.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as err:
        print("??", err)
        return str(err), 403

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

    EXCEL_FILE = {'buffer': request.files["file"], 'exception': ''}
    module_args = switcher.get(request.form["mock_selected"], None)()

    has_altered_sheet = insert_mock_data(EXCEL_FILE, module_args["column"], module_args["data"])

    return send_file(
        BytesIO(has_altered_sheet["buffer"].read()), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), has_altered_sheet["exception"] or None

@app.route("/capitalize_all", methods = ["POST"])
def get_capitalized_column():
    EXCEL_FILE = {'buffer': request.files["file"], 'exception': ''}
    
    has_altered_sheet = capitalize_all(EXCEL_FILE, request.form["columns_selected"].split(","))

    return send_file(
        BytesIO(has_altered_sheet["buffer"].read()), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), has_altered_sheet["exception"] or None

@app.route("/capitalize_firstletter", methods = ["POST"])
def get_capitalized_first():
    EXCEL_FILE = {'buffer': request.files["file"], 'exception': ''}

    has_altered_sheet = capitalize_firstLetter(EXCEL_FILE, request.form["columns_selected"].split(","))
    return send_file(
        BytesIO(has_altered_sheet["buffer"].read()), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), has_altered_sheet["exception"] or None

@app.route("/separate_addresses", methods = ["POST"])
def get_separated_addresses():
    EXCEL_FILE = {'buffer': request.files["file"], 'exception': ''}

    has_altered_sheet = separate_addresses(EXCEL_FILE, request.form["columns_selected"].split(","))
    return send_file(
        BytesIO(has_altered_sheet["buffer"].read()), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), has_altered_sheet["exception"] or None

@app.route("/separate_names", methods = ["POST"])
def get_separated_names():
    EXCEL_FILE = {'buffer': request.files["file"], 'exception': ''}

    has_altered_sheet = separate_names(EXCEL_FILE, request.form["columns_selected"].split(","))
    return send_file(
        BytesIO(has_altered_sheet["buffer"].read()), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), has_altered_sheet["exception"] or None
