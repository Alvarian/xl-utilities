from xl_utility import formatter
from xl_utility.formatter import *

from inspect import getmembers, isfunction
from io import BytesIO
from flask import Flask, request, render_template, send_file, session

app = Flask(__name__)
app.debug = True
app.secret_key = "development key"
app.config.update(
    SESSION_COOKIE_SECURE=True,
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE='Lax',
)

@app.route("/")
def index():
    # try:
    #     return send_file(
    #         "../demographics/main.xlsx", 
    #         download_name="python.jpg",
    #         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    #     )
    # except Exception as e:
    #     return str(e)
    omitted_functions = ["NamedTemporaryFile", "load_workbook", "get_column_letter"]
    feature_names = list(filter(lambda x: bool(x[0].split("_")[0]) and not x[0].split()[0] in omitted_functions, getmembers(formatter, isfunction)))
    return render_template("index.html", features=feature_names)

@app.route("/formatter/capitalize_all", methods = ["POST"])
def get_capitalized_column():
    EXCEL_FILE = BytesIO(request.files["file"].read())

    return send_file(
        BytesIO(capitalize_all(request.form["columns_selected"].split(","), EXCEL_FILE)["buffer"]), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/formatter/capitalize_firstletter", methods = ["POST"])
def get_capitalized_first():
    EXCEL_FILE = BytesIO(request.files["file"].read())

    return send_file(
        BytesIO(capitalize_firstLetter(request.form["columns_selected"].split(","), EXCEL_FILE)["buffer"]), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/formatter/separate_addresses", methods = ["POST"])
def get_separated_addresses():
    EXCEL_FILE = BytesIO(request.files["file"].read())

    return send_file(
        BytesIO(separate_addresses(request.form["columns_selected"].split(","), EXCEL_FILE)["buffer"]), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

@app.route("/formatter/separate_names", methods = ["POST"])
def get_separated_names():
    EXCEL_FILE = BytesIO(request.files["file"].read())

    # try:
    has_altered_sheet = separate_names(request.form["columns_selected"].split(","), EXCEL_FILE)

    return send_file(
        BytesIO(has_altered_sheet["buffer"]), 
        download_name="new_sheet.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    ), has_altered_sheet["exception"] or None
    # except Exception as error:
    #     print(error)

    #     return str(error), 403
