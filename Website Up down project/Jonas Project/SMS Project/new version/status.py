import pandas as pa
from mechanize import Browser
import pyodbc

pa.options.mode.chained_assignment = None
# Create connection
cnxn = pyodbc.connect(driver="{SQL Server}", server="192.168.10.221", database="Educube", uid="sa", pwd="globals@123",
                      autocommit=True)
cursor = cnxn.cursor()
cnxn.autocommit = True


def get_response_message(guid, un, pw, iterations=0):
    br = Browser()
    br.open("http://xmlapi.myvaluefirst.com/psms/").read()
    br.select_form(nr=1)
    br["data"] = '''<?xml version="1.0" encoding="ISO-8859-1"?>
<!DOCTYPE STATUSREQUEST SYSTEM "http://xmlapi.myvaluefirst.com/psms/dtd/requeststatusv12.dtd">
<STATUSREQUEST VER="1.2">
<USER USERNAME="''' + un + '''" PASSWORD="''' + pw + '''"/>
<GUID GUID="''' + guid + '''">
</GUID>
</STATUSREQUEST>'''
    response = br.submit()
    sTr = [s for s in str(response.read()).split() if s[0] == "E"]
    query = "Delivered" if int(sTr[0].split('"')[1]) == 8448 else "Failed"
    storedProc = "EXEC Educube.dbo.update_sms_status_from_excel_data @guid='" + guid + "', @status='" + query + \
                 "',@response='" + sTr[0].split('"')[1] + "'"
    cursor.execute(storedProc)
    #cnxn.commit()  # TODO Updating status to database
    try:
        return "Delivered" if int(sTr[0].split('"')[1]) == 8448 else "Failed"
    except IndexError:
        iterations += 1
        if iterations != 1:
            return get_response_message(guid, un, pw, iterations)
        else:
            return "Delivering"


if __name__ == '__main__':
    xl = pa.ExcelFile("SMS.xls")
    df = pa.read_excel(xl)
    for guid_index in range(len(df["GUID"])):
        try:
            df["Status"][guid_index] = get_response_message(df["GUID"][guid_index].split('"')[1],
                                                            df["User name"][guid_index],
                                                            df["Password"][guid_index])

        except:
            continue
    df.to_excel("SMS.xls", index=False)
