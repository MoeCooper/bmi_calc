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
        print(name, height, weight, bmi)
        save_to_excel(name, height, weight, bmi, data)
    return render_template("index.html", user_name=name, user_height=height, user_weight=weight, user_bmi=bmi)


def save_to_excel(name, height, weight, bmi, data):
    for name, weight, height, bmi in data:
        data.insert(name, height, weight, bmi)
    work_book = Workbook()
    work_sheet = work_book.active

    work_sheet.append(["Name", "Height", "Weight", "BMI"])
    work_sheet.append([name, height, weight, bmi])

    work_book.save("bmi_stats.xlsx")
    os.system("bmi_stats.xlsx")


def calculate_bmi(height, weight):
    return round(weight/pow(height/100, 2), 2)


app.run()
