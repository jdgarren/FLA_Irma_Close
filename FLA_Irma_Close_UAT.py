from py3270 import Emulator
from openpyxl import Workbook
from openpyxl import load_workbook
from datetime import datetime, date 
# use x3270 so you can see what is going on
wb1 = load_workbook(r'C:\Users\garren-james\Python Projects\FLORIDA Irma Closure PT\IRMA-Pending-Cases-County-UAT.xlsx')
ws1 = wb1['Pending IRMA Cases']
x = 1
y = input('Enter a number of records to process: ')
rn = input('Enter row number to start from: ')
rn = int(rn)
now = datetime.now()
today = now.strftime("%m%d%Y")

while x <= int(y): #specify number of times to run process
    em = Emulator(visible=True)
    
    startNum = 'A'+str(rn)
    caseNum = ws1[startNum].value
    caseNum = str(caseNum)
    em.connect('dcft.dcf.state.fl.us')
    em.wait_for_field()
    em.send_string('fla')
    em.send_enter()
    em.wait_for_field()
    em.fill_field(11, 36, 'fff0001', 7)
    em.fill_field(12, 36, 'DSNAP123', 8)
    em.send_enter()
    em.send_enter()
    em.send_enter()
    em.wait_for_field()
    em.fill_field(23, 13, 'AICI', 4)
    em.fill_field(23, 29, caseNum, 10)
    em.send_enter()
    em.wait_for_field()
    em.fill_field(23, 13, 'AABC', 4)
    em.fill_field(23, 29, caseNum, 10)
    em.send_enter()
    em.fill_field(7, 12, "01", 2)
    em.fill_field(7, 55, '01012019', 8)
    em.send_enter()
    if (em.string_found(24, 2, '075') ==True):
        em.fill_field(23, 13, 'AGPY', 4)
        em.exec_command(b'PF(14)')
        em.fill_field(7, 55, '01012019', 8)
        em.send_enter()
    else:
        pass
    if (em.string_found(1, 2, 'CUER') ==True):
        em.send_enter()
        em.fill_field(23, 13, 'AICI', 4)
        em.fill_field(23, 29, caseNum, 10)
        em.send_enter()
        em.move_to(12, 67)
        em.exec_command(b'Delete')
        em.exec_command(b'Delete')
        em.exec_command(b'Delete')
        em.exec_command(b'Delete')
        em.send_enter()
    else:
        pass
    em.fill_field(7, 29, '37081', 8)
    em.fill_field(7, 37, '37', 2)
    em.send_enter()

    if (em.string_found(1, 2, 'AGAR') == True):
        em.fill_field(6, 8, 'AE', 2)
        em.send_enter()
    else:
        pass
    if (em.string_found(24, 2, "MUST") == True):
        em.terminate()
        rn += 1
        print(rn)
        x += 1
    else:
        em.fill_field(8, 35, 'DE', 2)
        em.fill_field(9, 35, 'DE', 2)
        em.fill_field(10, 35, 'DE', 2)
        em.fill_field(10, 42, '374', 3)
        em.fill_field(10, 72, 'Y', 1)
        em.send_enter()
        em.wait_for_field()
        em.send_enter()
        em.fill_field(23, 13, 'CLRC', 4)
        em.fill_field(23, 29, caseNum, 10)
        em.send_enter()
        em.exec_command(b'PF(9)')
        em.fill_field(8, 18, "IRMA DSNAP CASE DENIED, CASE NOT PREVIOUSLY DISPOSED.", 53)
        em.send_enter()
        em.exec_command(b'PF(13)')
        em.send_enter()
        em.terminate()
        col2= 'B'+str(rn)
        ws1[col2] = "Done"
        wb1.save('IRMA-Pending-Cases-County-UAT.xlsx')
        rn += 1
        print(rn)
        x += 1

