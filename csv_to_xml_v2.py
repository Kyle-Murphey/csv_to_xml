import pandas as pd
import csv
import os
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import re

# splits the Excel workbook into separate CSV's and then calls generateXml() to convert to XML
def split(xfile):
    # split each sheet into separate, unformatted CSV's
    for key in xfile.keys():
        xfile[key].to_csv(f"{key}_temp.csv", index=False) #create temp CSV's of each sheet before formatting

        # create formatted CSV's and remove blank lines
        with open(f"{key}_temp.csv") as csv_temp_file:
            with open(f"{key}.csv", 'w', newline='') as csv_out_file:
                csv_writer = csv.writer(csv_out_file)
                for row in csv.reader(csv_temp_file, delimiter=','):
                    if any(row):
                        csv_writer.writerow(row)
        os.remove(f"{key}_temp.csv") #remove old temp CSV's

        # convert the CSV's to XML files
        generateXml(key)
    exit()


# converts the separate CSV's into separete XML files
def generateXml(key):
    nvpd_dictionary = {}
    with open(f"{key}.csv") as csv_file:
        csv_reader = csv.reader(csv_file, delimiter=',')
        temp_headers = next(csv_file).split(",")
        headers = []
        #pattern = re.compile('[^A-Za-z0-9_\-]')
        for header in temp_headers:
            temp = re.sub(" ", "-", header)
            #re.sub("\s", "", temp)
            #re.sub()
            headers.append(re.sub('[^A-Za-z0-9_\-]', '', temp))             #add headers below
            
        print(headers)
        columns = []

        with open(f"{key}.xml", 'w') as xmlfile:
            xmlfile.write("<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n")
            xmlfile.write(f"<{key}>\n".replace(" ", "-"))
            
            for row in csv_reader:
                if row[0] not in columns:
                    nvpd_dictionary[row[0]] = {row[1] : row[2]}
                    columns.append(row[0])
                    continue
                nvpd_dictionary[row[0]][row[1]] = row[2]

            for nvpd, item in nvpd_dictionary.items():
                xmlfile.write(f"<group>\n\t<NVPD>{nvpd}</NVPD>\n")
                for value in item:
                    xmlfile.write(f"\t\t<name>{value}</name>\n")
                    xmlfile.write(f"\t\t<value>{item[value]}</value>\n\n")
                xmlfile.write("</group>\n")

            xmlfile.write(f"</{key}>".replace(" ", "-"))
    os.remove(f"{key}.csv") #remove csv's
    print(f"{key}\tNVPDs: {len(columns)}") #debugging info - not necessary



def main():
    try:
        Tk().withdraw()
        filename = askopenfilename()
        xfile = pd.read_excel(filename, sheet_name=None)
        split(xfile)
    except Exception as e:
        print(e)
        exit()


if __name__ == "__main__":
    main()