from xl_utility import formatter, determinizer
from xl_utility.formatter import *
from xl_utility.determinizer import *

from inspect import getmembers, isfunction
from io import BytesIO
from flask import Flask, request, render_template, send_file, session

app = Flask(__name__)
app.debug = True
app.secret_key = "development key"


@app.route("/")
def index():
    omitted_functions = ["NamedTemporaryFile", "load_workbook", "get_column_letter", "generate_uuid"]

    formatter_features = list(filter(lambda x: bool(x[0].split("_")[0]) and not x[0].split()[0] in omitted_functions, getmembers(formatter, isfunction)))
    determinizer_features = list(filter(lambda x: bool(x[0].split("_")[0]) and not x[0].split()[0] in omitted_functions, getmembers(determinizer, isfunction)))

    feature_names = formatter_features + determinizer_features
    return render_template("index.html", features=feature_names)

@app.route("/capitalize_all", methods = ["POST"])
def get_capitalized_column():
    EXCEL_FILE = BytesIO(request.files["file"].read())

    has_altered_sheet = capitalize_all(request.form["columns_selected"].split(","), EXCEL_FILE)
    return send_file(
        BytesIO(capitalize_all(request.form["columns_selected"].split(","), EXCEL_FILE)["buffer"]), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), has_altered_sheet["exception"] or None

@app.route("/capitalize_firstletter", methods = ["POST"])
def get_capitalized_first():
    EXCEL_FILE = BytesIO(request.files["file"].read())

    has_altered_sheet = capitalize_firstLetter(request.form["columns_selected"].split(","), EXCEL_FILE)
    return send_file(
        BytesIO(capitalize_firstLetter(request.form["columns_selected"].split(","), EXCEL_FILE)["buffer"]), 
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
