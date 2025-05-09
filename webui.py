import flask
from flask import request, send_file
from main import conversion

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

app = flask.Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():

    # get the nessicary data to run the conversion. the user has submitted a file for us to process
    if request.method == "POST":

        # get the file name and sheet name from the form
        file = flask.request.files["file"]

        conversion(file)

        return send_file("Transformed.xlsx", download_name="Transformed.xlsx", mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", as_attachment=True)
    
    
    else:
        return flask.render_template("index.html")