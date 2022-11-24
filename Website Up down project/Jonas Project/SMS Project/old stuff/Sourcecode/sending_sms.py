from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import easyxf
from string import Template
from mechanize import Browser


def float_to_int(x):
    if isinstance(x, float):
        return int(x)
    return x


def send_sms(x, ):
    br = Browser()
    br.open("http://xmlapi.myvaluefirst.com/psms/").read()
    br.select_form(action="/psms/servlet/psms.Eservice2")
    br["data"] = x
    response = br.submit()
    responses = [s for s in str(response.read()).split() if s[0] == "G" or s[0] == "C"]
    if responses[0][0] == "G":
        return responses[0]
    else:
        return "Error" + responses[0]


if __name__ == '__main__':
    print("Starting")

    workbook = open_workbook("./SMS.xls")
    sheet = workbook.sheet_by_index(0)
    book_write = copy(workbook)
    sheet_write = book_write.get_sheet(0)
    try:
        response_col = [sheet.cell_value(0, i) for i in range(sheet.ncols)].index("GUID")
    except ValueError:
        print("Creating GUID Column")
        response_col = sheet.ncols
        sheet_write.write(0, response_col, "GUID")

    tF = open("./template.xml")
    string = tF.read()
    for i in range(1, sheet.nrows):
        wrF = open("./" + str(i) + ".xml", "w")
        liste = []
        for j in range(sheet.ncols - 1):
            body = str(float_to_int(sheet.cell_value(i, j)))
            liste.append("\"" + body + "\"")
        formatted = Template(string).substitute(fromNo=liste[0], to=liste[1], body=liste[2], un=liste[3], pw=liste[4])
        sheet_write.write(i, response_col, send_sms(formatted))
        wrF.write(formatted)
        wrF.close()
    del sheet
    del workbook
    book_write.save("./SMS.xls")
