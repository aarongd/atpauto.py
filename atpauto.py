import openpyxl

#Where information pulled from the Action Tracker, SLA Tracker, and cost org will go.
li = {
    "Candidate Name" : "",
    "Company" : "",
    "JITR" : "",
    "CSR" : "",
    "Labor Category" : "",
    "Level" : "",
    "Effective Date" : "",
    "Submitted Rate to CACI" : "",
    "SLA" : "",
    "CLIN" : "0820",
    "Resource ID" : "",
    "Cost Center" : "",
}

#the two inputs needed to put in manually
csr_input = input("CSR Number")
resource_id_input = input("Resource ID")

li["Resource ID"] = resource_id_input

#opens the Action Tracker workbook
n=0
wb = openpyxl.load_workbook(**workbook location here**)
sheets = wb.sheetnames
ws = wb[sheets[n]].active

#to find all rows within the CSR column in the Action tab
for row in ws.iter_rows("C"):
    for cell in row:
        if cell.value == csr_input:
            #saving all relevant values from this excel book to memory
            full_name = ws.cell(row=cell.row, column=2).value
            company = ws.cell(row=cell.row, column=2).value
            jitr = ws.cell(row=cell.row, column=2).value
            labor_category = ws.cell(row=cell.row, column=2).value
            level = ws.cell(row=cell.row, column=2).value
            start_date = ws.cell(row=cell.row, column=2).value
            rate2caci = ws.cell(row=cell.row, column=2).value
            #start adding all saved values from this excel book to the list
            li["Candidate Name"] = full_name
            li["Company] = company
            li["JITR"] = jitr
            li["CSR"] = csr_input
            li["Labor Category"] = labor_category
            li["Level"] = level
            li["Effective Date"] = start_date
            li["Submitted Rate to CACI"] = rate2caci
            return
        else:
            print("CSR value not found")

#opens the SLA tracker
wb2 = openpyxl.load_workbook(**workbook2 location here**)
sheets2 = wb.sheetnames
ws2 = wb2.active

#find all row relevant to the JITR # associated with the CSR
for row in ws.iter_rows("B"):
    for cell in row:
        if cell.value == jitr:
            sla = ws.cell(row=cell.row, column=2).value
            li["SLA"] = sla
        else:
            li["SLA"] = "Could not find SLA"

#Cost center is a constant unless for a few specific JITRs
if jitr == "1124" or "1125" or "1126" or "1158" or "1160" or "1166":
    li["Cost Center"] = 3373
else:
    li["Cost Center"] = 3393


