import openpyxl

#Where information pulled from the Action Tracker, SLA Tracker, and cost org will go.
li = {
    "JITR" : "",
    "CSR" : "",
    "Cost Center" : "",
}

#the two inputs needed to put in manually
csr_input = input("CSR Number: ")
#resource_id_input = input("Resource ID: ")

#li["Resource ID"] = resource_id_input

#opens the Action Tracker workbook
path = "C:\\Users\\Aaron\\Desktop\\action tracker.xlsx"
wb = openpyxl.load_workbook(path)
sheets = wb.sheetnames
ws = wb.active
#n=0
#ws = wb[sheets[n]].active

#to find all rows within the CSR column in the Action tab
for row in ws.iter_rows(min_row=2, min_col=3, max_col=3):
    for cell in row:
        if cell.value == csr_input:
            #saving all relevant values from this excel book to memory
            jitr = ws.cell(row=cell.row, column=2).value
            #start adding all saved values from this excel book to the list
            li["JITR"] = jitr
            li["CSR"] = csr_input
            print(jitr)
            # Cost center is a constant unless for a few specific JITRs
            if int(jitr) == (1124 or 1125 or 1126 or 1158 or 1160 or 1166):
                li["Cost Center"] = 3373
            else:
                li["Cost Center"] = 3393
        else:
            print("CSR value not found")

print(li)

