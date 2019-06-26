import loadingExcel as le
import generator as gg
import os, sys
import re,json

def addressAnalyzer(address_str):
    address={
        "fullAddress":"",
        "partialAddress":"",
        "city":"",
        "province":"",
        "abbreviation":"",
        "zipcode":""
    }
    try:
        with open('zip2prov.json', 'r') as f:
            data_zip = json.load(f)
    except Exception as result:
        print('Unexpected Error {}'.format(result))

    pattern_ZIP = r"[A-Z]\d[A-Z] \d[A-Z]\d"
    pattern_Province_A=r"\b[A-Z]{2}\b"

    str = "2390 47 Ave Calgary,AB  T2T 5W5"
    #address_str=str
    address["fullAddress"]=address_str

    match_zip=re.search(pattern_ZIP,address_str)
    if match_zip:
        address["zipcode"]=re.search(pattern_ZIP,address_str).group()
    else: print("no matching! ZIP CODE")

    match_Province_A = re.search(pattern_Province_A,address_str)
    if match_Province_A:
        address["abbreviation"]=match_Province_A.group()
    else:
        print("no matching! PROVINCE")
    if address["zipcode"] != "":
        for province in data_zip:
            if address["zipcode"][0] in province["ZFL"]:
                if province["Abbreviation"].upper() != address["abbreviation"].upper():
                    address["abbreviation"]=province["Abbreviation"].upper()
                address["province"]=province["Province"]
                for city in province["Cities"]:
                    if city.lower() in address_str.lower():
                        address["city"]=city
                        address["partialAddress"]=address_str[:address_str.find(city)]
    return address
#    print(address)

oriPath = os.getcwd()

path = r"C:\Users\fxiao\Documents\Change request #\ACAa002-ACAa006, ACAa015-ACAa016"
file ="Mission Trial ACAa002-ACAa006, ACAa015-ACAa016 LTE New Site RF CIQ.xlsx"
sitecode="ACAa002-ACAa006, ACAa015-ACAa016 "

dataset = le.load(path, file)
os.chdir(oriPath+"/docs")

dataset['Analyzed Address'] = [addressAnalyzer(a) for a in dataset['Property Address']]

gg.generate_datafill(dataset,sitecode,"CM")
gg.generate_datafill(dataset,sitecode,"SAS info")
gg.generate_datafill(dataset,sitecode,"unlock")
gg.generate_datafill(dataset,sitecode,"SCF")

'''
#this is to test dataset avalibility
print(len(dataset["Data"]))
data = dataset["Data"][0][10]
#this is to indentify a 'none' data
print(data)
print(type(data))
print(data==None) 
'''
