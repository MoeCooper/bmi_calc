import os
from flask import Flask, request, render_template
from openpyxl import Workbook

app = Flask(__name__)


@app.route('/', methods=["GET", "POST"])
def calculate():
    name = ""
    height = ""
    weight = ""
    bmi = ""
    data = []
    if request.method == "POST" and "username" in request.form:
        name = request.form.get("username")
        height = float(request.form.get("userheight"))
        weight = float(request.form.get("userweight"))
        bmi = calculate_bmi(height, weight)
        print(name, height, weight)
        for name, weight, height in data:
            data.insert(name, height, weight)
        work_book = Workbook()
        work_sheet = work_book.active

        work_sheet.append(["name", "height", "weight"])
        work_sheet.append([name, height, weight])

        work_book.save("bmi_stats.xlsx")
        os.system("bmi_stats.xlsx")
    return render_template("index.html", user_name=name, user_height=height, user_weight=weight, user_bmi=bmi)


def calculate_bmi(height, weight):
    return round(weight/pow(height/100, 2), 2)


app.run()