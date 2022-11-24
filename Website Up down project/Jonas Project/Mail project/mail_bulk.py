from xlrd import open_workbook
from xlutils.copy import copy
from xlwt import easyxf
from sendgrid import SendGridAPIClient, Attachment
from sendgrid.helpers.mail import Mail

if __name__ == '__main__':
    workbook = open_workbook("./bulkmailpy.xls")
    book_write = copy(workbook)
    sheet = workbook.sheet_by_index(0)
    sheet_write = book_write.get_sheet(0)
    try:
        response_col = [sheet.cell_value(0, i) for i in range(sheet.ncols)].index("Status")
    except ValueError:
        print("Creating Status Column")
        response_col = sheet.ncols
        sheet_write.write(0, response_col, "Status")

    dictlist = []
    responses = []
    for i in range(1, sheet.nrows):
        dictlist.append({})
        dictlist[i - 1] = {sheet.cell_value(0, j): sheet.cell_value(i, j) for j in range(sheet.ncols)
                           if sheet.cell_value(0, j) != ""}
        dictlist[i - 1]["To address"] = [k for k in dictlist[i - 1]["To address"].replace(" ", "").split(",") if k]
        dictlist[i - 1]["BCC Address"] = dictlist[i - 1]["BCC Address"].replace(" ", "").split(",")
        dictlist[i - 1]["CC address"] = dictlist[i - 1]["CC address"].replace(" ", "").split(",")
        mail = Mail(
            from_email="services@globalsinc.com",
            to_emails=dictlist[i - 1]["To address"],
            subject=dictlist[i - 1]["Subject"],
            html_content=dictlist[i - 1]["Message"],
        )
        for cc in dictlist[i - 1]["CC address"]:
            if cc != "":
                mail.add_cc(cc)
        for bcc in dictlist[i - 1]["BCC Address"]:
            if bcc != "":
                mail.add_cc(bcc)
        responses.append(
            SendGridAPIClient("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").send(mail)
        )
        sheet_write.write(i,response_col,responses[i - 1].status_code)
    book_write.save("./bulkmailpy.xls")
