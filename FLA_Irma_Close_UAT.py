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
    
    def logout ():
        em.fill_field(23, 13, 'smsf', 4)
        em.send_enter()
        em.send_enter()



    startNum = 'A'+str(rn)
    caseNum = ws1[startNum].value
    caseNum = str(caseNum)
    em.connect('dcft.dcf.state.fl.us')
    em.wait_for_field()
    em.send_string('floraccp')
    em.send_enter()
    em.wait_for_field()
    em.fill_field(11, 36, 'fz0383', 7)
    em.fill_field(12, 36, 'moxie123', 8)
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
    em.send_enter()
    em.fill_field(7, 29, '37081', 8)
    em.fill_field(7, 37, '37', 2)
    em.send_enter()
    em.fill_field(8, 35, 'DE', 2)
    em.fill_field(9, 35, 'DE', 2)
    em.fill_field(10, 35, 'DE', 2)
    em.fill_field(11, 35, 'DE', 2)
    em.fill_field(12, 35, 'DE', 2)
    em.fill_field(13, 35, 'DE', 2)
    em.fill_field(13, 42, '374', 3)
    em.fill_field(13, 72, 'Y', 1)
    em.send_enter()
    em.wait_for_field()
    em.send_enter()
    em.fill_field(23, 13, 'CLRC', 4)
    em.fill_field(23, 29, caseNum, 10)
    em.send_enter()
    em.exec_command(b'PF(9)')
    em.fill_field(8, 18, "IRMA DSNAP CASE CLOSED, NOT IN DISASTER AREA.", 45)
    em.send_enter()
    logout()
    em.terminate()
    rn += 1
    print(rn)
    x += 1

