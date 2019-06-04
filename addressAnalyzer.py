import re,json

address={
    "fullAddress":"",
    "partialAddress":"",
    "city":"",
    "province":"",
    "abbreviation":"",
    "zipcode":""
}
def main(address_str):
    try:
        with open('zip2prov.json', 'r') as f:
            data_zip = json.load(f)
    except Exception as result:
        print('检测出异常{}'.format(result))

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
                            
                            
                            
    print(address)
main("6201 Country Club Rd NW, Edmonton, AB T6M 2J6")
print()
main("2830 Bentall Street, Vancouver, BC, V5M 4H4")
print()
main("5720 Silver Springs Blvd NW, Calgary, AB T3B 4N7")

"""
if __name__=="__main__":
    main(address_str)
"""