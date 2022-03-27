from datetime import datetime
import re
import codecs
from openpyxl import Workbook
phone_list = []

now = datetime.now()
now = now.strftime("%Y-%m-%d %H:%M:%S").replace(' ', '_')

ContactList = []

def filter_messages ( message ):
    if 'attached' in message:
        message = message.split('attached:', 2)[1].replace('>', '')
        
        if not message.endswith(".vcf"): 
            return False
        message = message.strip()
        message = 'whatsappfile/{}'.format( message )
        return message
    return False
    

def saveContact ( file_name ):
    with codecs.open(file_name, 'r', encoding="utf-8") as data:
        lines = data.read().splitlines()
        for line in lines:
            if 'FN:' in line:
                full_name = line.split(':', 1)[1]
            if 'TEL' in line:
                phone = line.split(':', 2)[1]
                phone_list.append({
                    'full_name': full_name,
                    'phone': phone
                })
    
def saveToFile ():
    row = 1
    wb = Workbook()
    wb['Sheet'].title = "List of shared contacts"
    sh1 = wb.active
    for i in range(len(phone_list) + 1):
        sh1['A{}'.format(i+1)].value = phone_list[i-1].get('full_name')
        sh1['B{}'.format(i+1)].value = phone_list[i-1].get('phone')
    wb.save("shared_people.xlsx")
    
with open('whatsappfile\_chat.txt', 'r', encoding='utf8') as file:
    lines = file.read().splitlines()
    
    for line in lines:
        contactFile = filter_messages( line )
        if contactFile:
            saveContact( contactFile )
    saveToFile()