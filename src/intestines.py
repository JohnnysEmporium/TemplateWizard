from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from os import environ, getcwd
import docx, pyperclip, re, os, getpass, sys, subprocess


userName = getpass.getuser() 

def choose_file():
    arr = ["output.docx"]
    for file in os.listdir(os.path.join(os.getcwd(), "MSG")):
        if file.endswith(".msg"):
            new_name = file
            if "INC" in file and len(file) > 20:
                new_name = "P" + [s for s in file.split() if s.isdigit()][0] + "_" + file[file.find('INC'):file.find('INC')+10] + ".msg"
#                 new_name = "2.msg"
                os.rename(os.path.join(os.getcwd(), "MSG", file), os.path.join(os.getcwd(), "MSG", new_name))
            arr.append(new_name)
    
    print("\nSelect a file to work with:\n---------------")
    for i, name in enumerate(arr):
        print(str(i+1) + ". " + name)
    
    print("---------------")            
    fnameNo = input()
    
    if int(fnameNo) < 1 or int(fnameNo) > len(arr):
        print('\n!!!---You must input a number between 1 and ' + str(len(arr)) + '---!!!\n')
        choose_file()
    else:
        return arr[int(fnameNo) - 1]

# Takes care of properly displaying User Name
def getUserName():
    x = userName.split('.')
    name = x[0].capitalize()
    surname = x[1].capitalize()
    return (name + ' ' + surname)

# Takes care of saving files
def save_file(doc, prio, incNo, stat, fname = 'output.msg'):
    
    for file in os.listdir():
        if file.endswith(".msg"):
            os.rename(file, "output.msg")
    
    finalTouch(doc.tables[0])
        
#     try:
    doc.save('output.docx')
    print('\n---Filled template saved in Template Master source folder in "output.docx"---\n')
    os.system('MSG\out.vbs ' + fname + " " + prio + " " + incNo + " " + stat)
#     except PermissionError:
#         print('\n!!!---File in use, close output.docx and press ENTER to continue, type "stop" to cancel---!!!\n')
#         x = input()
#         if x == 'stop':
#             pass
#         else:
#             save_file(doc, prio, incNo, stat, fname)
        
        
# Makes sure that the text is correctly formatted 
def finalTouch(tab):
    try:
        tab.rows[1].cells[0].paragraphs[0].runs[0].font.bold = True
        tab.rows[1].cells[0].paragraphs[0].runs[0].font.underline = True
        tab.rows[1].cells[0].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    except:
        pass

    for row in tab.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    font = run.font
                    font.size= Pt(12)

def paster():
    
# Searches for value given in the argument in the description. When found returns phrase until \n. When not founds returns -1
    def impacts():
        serviceImpact = input('''
Select Service Impact:
---------------
1. Messaging
2. Application
3. Network
4. Server
5. Workstation
6. Printer
7. Industrial Device
---------------        
''')

        if serviceImpact == '1':
            serviceImpact = 'Messaging'
            ciImpact = 'Messaging'            
        elif serviceImpact == '2':
            serviceImpact = 'Application'
            ciImpact = 'Application'
        elif serviceImpact == '3':
            serviceImpact = 'Network'
            ciImpact = 'Infrastructure'
        elif serviceImpact == '4':
            serviceImpact = 'Server'
            ciImpact = 'Infrastructure'
        elif serviceImpact == '5':
            serviceImpact = 'Workstation'
            ciImpact = 'Infrastructure'
        elif serviceImpact == '6':
            serviceImpact = 'Printer'
            ciImpact = 'Infrastructure'
        elif serviceImpact == '7':
            serviceImpact = 'Industrial Device'
            ciImpact = 'Industrial Device'
        else:
            print('\n!!!---You must input a number between 1 and 7---!!!\n')
            
        businessImpact = input('''
---------------
Select Business Impact
1. Production                   (Production line stopped)
2. Financial                    (Financial loss)
3. Shipping                     (train or truck blocked)
4. Safety                       (Employee safety)
5. Security                     (data or access security)
6. Multiple users from different locations are no able perform daily work
---------------
''')
    
        if businessImpact == '1':
            businessImpact = 'Production'
        elif businessImpact == '2':
            businessImpact = 'Financial'
        elif businessImpact == '3':
            businessImpact = 'Shipping'
        elif businessImpact == '4':
            businessImpact = 'Safety'
        elif businessImpact == '5':
            businessImpact = 'Security'
        elif businessImpact == '6':
            businessImpact = 'Multiple users from different locations are no able perform daily work'
            
        return [serviceImpact, ciImpact, businessImpact]
    
    def parser(x):
        n = description.find(x)
        if x == 'ISSUE DESCRIPTION:':

            if description.find('[ENG]') != -1:
                n = description.find('[ENG]')
                return (description[n+5:] if n != -1 else -1)
            else:
                m = n+30            
                return (description[n:n+description[m:].find('\n')].split(':', 1)[1] if n != -1 else -1)
        else:
            return (description[n:n+description[n:].find('\n')].split(':', 1)[1] if n != -1 else -1)
        
# Assigning data to variables
    doc = docx.Document('output.docx')
    table = doc.tables[0]
    doc = docx.Document('template\\template.docx')
    table = doc.tables[0]
    data = pyperclip.paste()
    data = data.split('/nextEl,')
    if len(data) == 10:
        incNo = data[0] 
        incStatus = data[1] 
        incPrio = data[2] 
        summary = data[3]
        description = data[4]
        RG = data[5]
        startDate = data[6]
        latestDate = data[7]
        latestUpdate = data[8]
        desc = parser('ISSUE DESCRIPTION:')
        location = parser('LOCATION')
        impact = impacts()
          
# Filling the table with scrapped values
        table.cell(1,0).text = '\n' + 'P' + incPrio + ' ' + incNo + ' Incident Initial Notification' + '\n'
        table.cell(2,1).text = incNo
        table.cell(3,1).text = startDate 
        table.cell(4,1).text = 'P' + incPrio
        table.cell(4,3).text = incStatus
        table.cell(5,1).text = summary
        table.cell(6,1).text = (desc.strip() if desc != -1 else '')
        table.cell(7,1).text = impact[0]
        table.cell(7,3).text = impact[1]
        table.cell(8,1).text = impact[2]
        table.cell(10,1).text = (location if location != -1 else '')
        table.cell(11,1).text = getUserName()
        table.cell(11,3).text = RG
        table.cell(13,1).text = latestDate + ' - ' + latestUpdate
        table.cell(14,1).text = ('30 minutes' if incPrio == '1' else 'Upon Resolution')
         
        save_file(doc, incPrio, incNo, "INITIAL")
    else:
        print('\n!!!---Invalid data format, press ALT+5 in SNow and try again---!!!\n')
    
def colors(x):
    
# Fills the template with choosen colors, can also add latest work-notes update
    def filling(val):
        fill1 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill2 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill3 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill4 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill5 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill6 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill7 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill8 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill9 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill10 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill11 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill12 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill13 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill14 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill15 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
        fill16 = parse_xml(r'<w:shd {} w:fill="{}"/>'.format(nsdecls('w'), val))
            
        table.cell(2,0)._tc.get_or_add_tcPr().append(fill1)
        table.cell(3,0)._tc.get_or_add_tcPr().append(fill2)
        table.cell(4,0)._tc.get_or_add_tcPr().append(fill3)
        table.cell(4,2)._tc.get_or_add_tcPr().append(fill4)
        table.cell(5,0)._tc.get_or_add_tcPr().append(fill5)
        table.cell(6,0)._tc.get_or_add_tcPr().append(fill6)
        table.cell(7,0)._tc.get_or_add_tcPr().append(fill7)
        table.cell(7,2)._tc.get_or_add_tcPr().append(fill8)
        table.cell(8,0)._tc.get_or_add_tcPr().append(fill9)
        table.cell(9,0)._tc.get_or_add_tcPr().append(fill10)
        table.cell(10,0)._tc.get_or_add_tcPr().append(fill11)
        table.cell(11,0)._tc.get_or_add_tcPr().append(fill12)
        table.cell(11,2)._tc.get_or_add_tcPr().append(fill13)
        table.cell(12,0)._tc.get_or_add_tcPr().append(fill14)
        table.cell(13,0)._tc.get_or_add_tcPr().append(fill15)        
        table.cell(14,0)._tc.get_or_add_tcPr().append(fill16)
    
    chFile = choose_file()
    
    if chFile == "output.docx":
        doc = docx.Document('output.docx')
        table = doc.tables[0]
#         latest_update = pyperclip.paste()
#         latest_update = latest_update.split('/nextEl,')
    else: 
        os.system('MSG\in.vbs ' + chFile)
        doc = docx.Document('MSG/temp.docx')
        table = doc.tables[0]
        
    incNo = table.cell(2,1).text
    incPrio = table.cell(4,1).text[1]
 
    if x == '1':
        val = 'FF0000'
        filling(val)
        save_file(doc, incPrio, incNo, "INITIAL", chFile)
    elif x == '2':
        val = 'FFC000'
        table.cell(1,0).text = table.cell(1,0).text.replace('Initial', 'Update')
        filling(val)
        save_file(doc, incPrio, incNo, "UPDATE", chFile)
    elif x == '3':
        val = '00B050'
        table.cell(1,0).text = table.cell(1,0).text.replace('Initial', 'Final')
        filling(val)
        save_file(doc, incPrio, incNo, "FINAL", chFile)
    elif x == '4':
        if len(latest_update) == 3:
            previous_update = table.cell(12,1).text
            table.cell(13,1).text = latest_update[0] + ' - ' + latest_update[1] + '\n\n' + previous_update
            save_file(doc, incNo, incPrio, "UPDATE", chFile)
        else: 
            print('\n!!!---Invalid data format, press ALT+6 in SNow and try again---!!!\n')
            
    os.remove(os.path.join(os.getcwd(), "MSG", "temp.docx"))
    