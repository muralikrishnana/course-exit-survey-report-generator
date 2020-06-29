import os
import json
import csv
import sys
from docxtpl import DocxTemplate

REPORT_PATH = 'survey-reports'
QUESTION_PATH = 'sample_questions.json'
DATA_PATH = 'sample_data.csv'

if not os.path.exists(REPORT_PATH):
    os.makedirs(REPORT_PATH)

tickMark = "✔"

config = {
    "course_code": "",
    "course_name": "",
    "faculty_name": "",
    "course_session": "",
    "semester": {
        "roman": "",
        "text": ""
    },
    "department": "",
}


def getQuestions():
    with open(QUESTION_PATH) as chunk:
        data = json.load(chunk)

    return data


def getData():
    data = []

    with open(DATA_PATH) as file:
        reader = csv.DictReader(file)

        for row in reader:
            data.append(row)

    return data


def preProcessData(data, questions):

    for datum in data:
        datum["course_code"] = config["course_code"]
        datum["course_name"] = config["course_name"]
        datum["faculty_name"] = config["faculty_name"]
        datum["course_session"] = config["course_session"]
        datum["semester"] = config["semester"]
        datum["department"] = config["department"]

        datum["name_of_student"] = datum["name"].upper()
        datum["roll_no_of_student"] = datum["roll_no"]

        datum["questions"] = []

        for i in range(0, len(questions)):
            question = questions[i]
            q = {
                "no": i+1,
                "text": question,
                "vh": tickMark if datum[question] == "vh" else "",
                "h": tickMark if datum[question] == "h" else "",
                "m": tickMark if datum[question] == "m" else "",
                "l": tickMark if datum[question] == "l" else "",
                "vl": tickMark if datum[question] == "vl" else ""
            }

            datum["questions"].append(q)

    return data


def generate(pathToTemplate):
    template = DocxTemplate(pathToTemplate)
    data = getData()
    questions = getQuestions()

    data = preProcessData(data, questions)

    print("\nPlease wait while the reports are generated...")

    for datum in data:
        template.render(datum)

        student_name = datum["name"].upper()
        student_name = student_name.replace(".", "")
        student_name = student_name.split(" ")
        student_name = "".join(student_name)

        fname = "{student_name}_{student_roll_no}.docx".format(
            student_name=student_name, student_roll_no=datum["roll_no"])

        filePath = r"survey-reports/" + fname

        template.save(filePath)

    print("\nReports generated. Check " + REPORT_PATH + " folder.")


if __name__ == "__main__":
    # arg 1 -> questions
    # arg 2 -> csv
    if len(sys.argv) > 2:
        if (not os.path.isfile(sys.argv[1])):
            print("File with questions required as first argument. Exiting.")
            sys.exit()

        if (not os.path.isfile(sys.argv[2])):
            print("File with csv required as first argument. Exiting.")
            sys.exit()

        QUESTION_PATH = sys.argv[1]
        DATA_PATH = sys.argv[2]

    print("Course Report Generator")

    config["course_code"] = input("\nEnter course code: ").upper()
    config["course_name"] = input("Enter course name: ")
    config["faculty_name"] = input("Enter faculty name: ")
    config["course_session"] = "Odd" if int(
        input("Enter course session (1 - Odd, 2 - Even): ")) == 1 else "Even"
    config["semester"]["roman"] = input("Enter semester roman: ").upper()
    config["semester"]["text"] = input("Enter semester text: ").upper()
    config["department"] = input("Enter department: ").upper()

    generate("./template.docx")
