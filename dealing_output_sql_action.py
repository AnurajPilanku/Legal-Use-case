#ANURAJ PILANKU
import sys
import openpyxl
import time
time.sleep(2)
lastoutput=sys.argv[1]
excelpath=sys.argv[2]
if lastoutput=="No records found!!":
    wb=openpyxl.Workbook()
    ws=wb.active
    wb.save(excelpath)
    print("success")
elif  "placed".strip() in lastoutput.split():
    print("success")
elif lastoutput.lower().strip()=="Workflow Returned Empty Output".lower():
    print("success")
elif "executed Successfully".lower().strip() in lastoutput.lower().strip():
    print("success")
else:
    print("failure")
    #="The output file is placed in host server.File location :- \acdev01\3M_CAC\AOMS_user_role\access.xlsx

