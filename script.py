#!/usr/bin/env python
# -*- coding: utf-8 -*-

import xlrd
import codecs
import openpyxl
import re
import os

filevalue = os.getcwd() + "/Desktop/script/"

coma = re.compile(",")
enie = re.compile("ñ")

def InsertFromods(filename):
    """Make an LDAP entry from an Excel file."""
    wb = openpyxl.load_workbook(filename)
    sh = wb.worksheets[0]

    # Open the output file for writing
    with codecs.open("entries.ldif", "a+", "utf-8") as filen:
        val = 2435  # The last uidNumber attribute in the LDAP directory

        # Iterate through the rows
        for rownum in range(1, sh.max_row + 1):
            row = [cell.value for cell in sh[rownum]]
            
            if len(row) > 1 and row[1] is not None:  # Ensure the second cell exists and is not None
                value = row[1].lower()  # Access the value in the second cell and convert it to lowercase
                value = value.replace(u"ñ", u"n")  # Replace the character 'ñ' with 'n' in the value

                fsfriendlyname = row[2] if len(row) > 2 else ""
                mail = row[3] if len(row) > 3 else ""
                phonenumber = str(row[4]) if len(row) > 4 else ""
                title = row[5] if len(row) > 5 else ""
                userPassword = row[6] if len(row) > 6 else ""

                # Try to organize the second column values
                if coma.search(value):
                    ape = value.split(",")[0].strip()  # If there is a comma, split by comma and get the first part
                    nom = value.split(",")[-1].strip()  # Get the last part after the comma
                else:
                    ape = value.split()[0].strip()  # If there is no comma, split by spaces and get the first part
                    nom = value.split()[-1].strip()  # Get the last part after spaces

                # Write formatted data to the file
                filen.write("dn: uid=%s%s,ou=User,ou=Dma2ljbxm,o=APAC,dc=fiserv,dc=com\n" % (nom, ape))
                filen.write("objectClass: fsPerson\n")
                filen.write("objectClass: fsUser\n")
                filen.write("objectClass: inetOrgPerson\n")
                filen.write("objectClass: person\n")
                filen.write("objectClass: top\n")
                filen.write("cn: %s %s\n" % (nom, ape))
                filen.write("mail: %s\n" % mail)
                filen.write("phonenumber: %s\n" % phonenumber)
                filen.write("title: %s\n" % title)
                filen.write("userPassword: %s\n" % userPassword)
                filen.write("sn: %s\n" % ape)
                filen.write("fsowner: UAID-TESTING\n")
                filen.write("fsfriendlyname %s\n" % fsfriendlyname)

                val += 1
                filen.write("uid: %s%s\n" % (nom, ape))
                # filen.write("uidNumber: %s\n" % str(val))
                # filen.write("gidNumber: %s\n" % str(val))
                # filen.write("homeDirectory: /homedirs/%s%s\n" % (nom, ape))
                filen.write("\n")

if __name__ == '__main__':
    InsertFromods("test.xlsx")
