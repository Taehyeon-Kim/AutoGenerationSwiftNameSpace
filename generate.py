import os
import sys
import urllib
import importlib
import csv


reload(sys)

gdoc_id = "1WMnHk1brX7CR34hCN88cDnTDc4UgTAKh-qegwnPIqCU/edit#gid=0"


def get_gdoc_information():
    download_path = sys.argv[1]
    try:
        csv_file = export_xlsx_from_sheet(gdoc_id)
        string_list = get_strings_from_csv(download_path, csv_file)
        for string in string_list.items():
            write_strings(string[0], string[1])
            
    except Exception as e:
        print(":::::::::::::ERROR:::::::::::::")
        print(e)


def export_xlsx_from_sheet(gdoc_id, download_path=None, ):
    print("Downloading the XLSX file with id %s" % gdoc_id)

    resource = gdoc_id.split('/')[0]
    resource_id = 'spreadsheet:' + resource

    if download_path is None:
        download_path = os.path.abspath(os.path.dirname(__file__))

    file_name = os.path.join(download_path, '%s.csv' % (resource))

    print('download_path : %s' % download_path)
    print('Downloading spreadsheet to %s' % file_name)

    url = 'https://docs.google.com/spreadsheet/ccc?key=%s&output=csv' % (
        resource)
    urllib.urlretrieve(url, file_name)

    print("Download Completed!")

    return file_name

def get_strings_from_csv(savepath, file_name):
    print("read CSV file : %s" % file_name)

    source_csv = open(file_name, "r")
    csv_reader = csv.reader(source_csv)
    header = next(csv_reader)
    categories = ["Text", "Color", "Image", "Storyboard", "Xib"]
    categoryIndex = []
    string_list = {}
    next(csv_reader)
    
    for category in categories:
        string_list[category] = []
        categoryIndex.append(header.index(category))


    for row in csv_reader:
        for index in categoryIndex:
            key = row[index]
            value = row[index + 1]

            dict_string = {
                "key": key,
                "value": value
            }

            if dict_string["key"] != "" and dict_string["value"] != "":
                string_list[categories[index//2]].append(dict_string)

    source_csv.close()
    os.remove(file_name)

    return string_list
    

def write_strings(filename, string_list):
    swift_file = open(filename+".swift", "w")

    swift_file.write("import UIKit\n\n")
    swift_file.write("enum "+ filename + " {\n")

    for item in string_list:
        if filename == "Image":
            swift_file.write("\tstatic let " + item["key"] + " = " + "UIImage(named: \"" + item["value"] + "\")\n")
        elif filename == "Color":
            swift_file.write("\tstatic let " + item["key"] + " = " + "UIColor(named: \"" + item["value"] + "\")\n")
        else:
            swift_file.write("\tstatic let " + item["key"] + " = " + "\"" + item["value"] + "\"\n")

    swift_file.write("}")
    swift_file.close()



if __name__ == '__main__':
    get_gdoc_information()
