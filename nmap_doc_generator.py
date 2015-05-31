#!/usr/bin/env python
# Author: Chris Duffy
# Date: May 2015
# Name: nmap_doc_generator.py
# Purpose: A script that takes data from parsed nmap XML files and writes it into XLSX files
try:
    import docx
    from docx.shared import Inches
except:
    sys.exit("[!] Install the docx writer library as root or through sudo: pip install python-docx")
try:
    import xlsxwriter
except:
    sys.exit("[!] Install the xlsx writer library as root or through sudo: pip install xlsxwriter")

class Nmap_doc_generator():
    def __init__(self, verbose, hosts_dict, filename):
        self.hosts_dict = hosts_dict
        self.filename = filename
        self.verbose = verbose
        try:
            self.run(verbose, self.hosts_dict, self.filename)
        except Exception as e:
            print(e)

    def run(self, verbose, hosts_dict, filename):
        print ("") #DEBUG
        # Run the appropriate module
        if verbose > 0:
            print ("[*] Building %s.xlsx") % (filename)
            self.generateXSLX(self.verbose, self.hosts_dict, self.filename)

    def generateXSLX(self, verbose, hosts_dict, filename):
        if not filename:
            filename = "%s.xlsx" % (xml)
        else:
            filename = "%s.xlsx" % (filename)
        workbook = xlsxwriter.Workbook(filename)
        # Row one formatting
        format1 = workbook.add_format({'bold': True})
        format1.set_bg_color('#538DD5')
        # Even row formatting
        format2 = workbook.add_format({'text_wrap': True})
        format2.set_align('left')
        format2.set_align('top')
        format2.set_border(1)
        # Odd row formatting
        format3 = workbook.add_format({'text_wrap': True})
        format3.set_align('left')
        format3.set_align('top')
        format3.set_bg_color('#C5D9F1')
        format3.set_border(1)
        if verbose > 0:
            print ("[*] Creating Workbook: %s") % (filename)

        # Generate Worksheet 1
        worksheet = workbook.add_worksheet("All Ports")
        # Column width for worksheet 1
        worksheet.set_column(0, 0, 20)
        worksheet.set_column(1, 1, 17)
        worksheet.set_column(2, 2, 22)
        worksheet.set_column(3, 3, 8)
        worksheet.set_column(4, 4, 26)
        worksheet.set_column(6, 6, 13)
        worksheet.set_column(6, 6, 12)
        # Define starting location for Worksheet one
        row = 1
        col = 0
        # Generate Row 1 for worksheet one
        worksheet.write('A1', "Hostname", format1)
        worksheet.write('B1', "Address", format1)
        worksheet.write('C1', "Hardware Address", format1)
        worksheet.write('D1', "Port", format1)
        worksheet.write('E1', "Service Name", format1)
        worksheet.write('F1', "Protocol", format1)
        worksheet.write('G1', "Port State", format1)
        worksheet.autofilter('A1:G1')
        # Populate Worksheet 1
        for key, value in hosts_dict.items():
            try:
                hostname = value[0]
                address = value[1]
                protocol = value[2]
                port = value[3]
                service_name = value[4]
                hwaddress = value[5]
                state = value[6]
            except:
                if verbose > 3:
                    print("[!] An error occurred parsing host ID: %s for Worksheet 1") % (key)
            try:
                if row % 2 != 0:
                    temp_format = format2
                else:
                    temp_format = format3
                worksheet.write(row, col,     hostname, temp_format)
                worksheet.write(row, col + 1, address, temp_format)
                worksheet.write(row, col + 2, hwaddress, temp_format)
                worksheet.write(row, col + 3, port, temp_format)
                worksheet.write(row, col + 4, service_name, temp_format)
                worksheet.write(row, col + 5, protocol, temp_format)
                worksheet.write(row, col + 6, state, temp_format)
                row += 1
            except:
                if verbose > 3:
                    print("[!] An error occurred writing data for Worksheet 1")
        try:
            workbook.close()
        except:
            sys.exit("[!] Permission to write to the file or location provided was denied")
