#!/usr/bin/env python
#-*- coding: utf-8 -*-

import xlrd
import codecs
import openpyxl
import re
import os
filevalue = os.getcwd()+"/Desktop/script/"

coma = re.compile(",")
enie = re.compile("ñ")

#----------------------------------------------------------------------
def InsertFromods(filename):
    # print(filename)
    """make a ldap entry from a ods file"""
    # wb = xlrd.open_workbook(filename)
    wb = openpyxl.load_workbook(filename)
   
    # sh = wb.sheet_by_index(0)
    sh = wb.worksheets[0]
    

    #for every entry we make a ldif
    filen = codecs.open("entries.ldif", "a+", "utf-8")
    #the last uidNumber attribute in the ldap directory
    val = 2435
   
    for rownum in range(1, sh.max_row + 1):
        # row = sh.row_values(rownum)
        row = [cell.value for cell in sh[rownum]]
        # print(row)
        if len(row) > 1 and row[1] is not None:  # Ensure the second cell exists and is not None
            value = row[1].lower()  # Access the value in the second cell and convert it to lowercase
            fsfriendlyname=row[2]
            mail=row[3]
            phonenumber=str(row[4])
            title=row[5]
            userPassword=row[6]
            # print(fsfriendlyname)
        value = value.replace(u"ñ", u"n")  # Replace the character 'ñ' with 'n' in the value

        # Try to organize the second column values
        if coma.search(value):
            ape = value.split(",")[0].strip()  # If there is a comma, split by comma and get the first part
            # print(ape)
            nom = value.split(",")[-1].strip()  # Get the last part after the comma
            # print(nom)
        else:
            ape = value.split()[0].strip()  # If there is no comma, split by spaces and get the first part
            print(value.split()[0].strip())
            nom = value.split()[-1].strip()  # Get the last part after spaces
            # print(nom)

          
        #pass
        filen.write("dn: uid=%s%s,ou=User,ou=Dma2ljbxm,o=APAC,dc=fiserv,dc=com\n" % (nom, ape))
        filen.write("objectClass: fsPerson\n")
        filen.write("objectClass: fsUser\n")
        filen.write("objectClass: inetOrgPerson\n")
        filen.write("objectClass: person\n")
        filen.write("objectClass: top\n")
        filen.write("cn: %s %s\n"  % (nom, ape))
        filen.write("mail: %s %s\n"  % (nom, ape))
        filen.write("phonenumber : %s %s\n"  % phonenumber)
        filen.write("title: %s %s\n"  % title)
        filen.write("userPassword: %s %s\n"  % userPassword)
        filen.write("sn: %s\n"  % ape)
        filen.write("fsowner: UAID-TESTING")
        filen.write("NumeroDocumentoIdentidad: %s\n" % row[2])           
        
        val = val + 1
        filen.write("uid: %s%s\n" % (nom, ape))                       
        # filen.write("uidNumber: %s\n"  % str(val))
        # filen.write("gidNumber: %s\n"  % str(val))
        # filen.write("homeDirectory: /homedirs/%s%s\n" % (nom, ape))

        filen.write("\n")
        filen.write("\n")           
    filen.close()
   
#----------------------------------------------------------------------

if __name__=='__main__':

    InsertFromods("test.xlsx")
    
    # backup 3