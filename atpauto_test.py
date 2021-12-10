import openpyxl
import win32com.client as win32

#Where information pulled from the Action Tracker, SLA Tracker, and cost org will go.
li = {
    "Candidate Name" : "Luke Skywalker",
    "Company" : "Star Wars",
    "JITR" : "111",
    "CSR" : "2021-98765",
    "Labor Category" : "Software Developer",
    "Level" : "SME",
    "Effective Date" : "11/21/21",
    "Submitted Rate to CACI" : "121.42",
    "SLA" : "fda3fasdf",
    "CLIN" : "0820",
    "Resource ID" : "234823",
    "Cost Center" : "3393",
}


#to send the email
outlook = win32.Dispatch('outlook.application')
mail = outlook.CreateItem(0)
#mail.To = 'aarongd1995@gmail.com'
mail.Subject = 'ATP Notification - ' + li['Candidate Name']
print(mail.Subject)
mail.Body = 'Kim,\n\n' \
            'Plase find the ATP Notification below:\n\n'\
            'Candidate Name: ' + li['Candidate Name']+ '\n'\
            'Postion #: JITR ' + li['JITR'] + ' / ' + li['CSR']+'\n'\
            'Labor Category: ' + li['Labor Category'] + '\n'\
            'Level: ' + li['Level'] + '\n\n'\
            'Effective Date: ' + li['Effective Date']+'\n'\
            'Submitted Rate to CACI: $' + li['Submitted Rate to CACI']+'\n'\
            'SLA: ' + li['SLA']+'\n'\
            'Resource ID: ' + li['Resource ID']+'\n'\
            'CLIN: ' + li['CLIN']+'\n'\
            '---\n' \
            'CACI Internal/FYI Kimberly\n' \
            'Resource ID: ' + li['Resource ID']+'\n'\
            'Cost Center: ' + li['Cost Center']+'\n'\
            '--------------------------------------------------------------------\n\n' \
            'Regards,\nAaron Davis | AGDS Lead Staffing Coordinator\nITDAS.PMO@CACI.com\nIntel Applications Services\n1540 Conference Center Drive | Suite 100 | Chantilly, Va 20151\n' \
            'Office: 703.667.9197 | Cell: 202.329.3537\nAaron.Davis@CACI.com | ww.caci.com'

print(mail.Body)
#mail.HTMLBody = '<h2>HTML Message body</h2>' #this field is optional

#mail.Send()