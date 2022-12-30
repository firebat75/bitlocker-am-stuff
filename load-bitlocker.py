from openpyxl import load_workbook
import requests
import json


# import keys from exo spreadsheet to variable data
data = {}
exo_data = load_workbook(filename = "exo-data.xlsx")
sheet = exo_data.active
i = 1
while sheet[f"a{i}"].value:
    data[sheet[f"f{i}"].value] = sheet[f"c{i}"].value
    print(sheet[f"f{i}"].value, sheet[f"c{i}"].value)
    i += 1



def getWorkstations(x, y=0):
    url = "http://10.3.1.220:443/api/v3/workstations"
    headers = {"authtoken":"7C6702E0-321E-417D-BF35-8DD79B1BE8FD"}
    list_info = {
        "list_info": {
            "row_count": x,
            "start_index": y,
            "sort_field": "id",
            "sort_order": "asc",
            "get_total_count": True
        }
    }
    params = {'input_data': json.dumps(list_info)}
    response = requests.get(url,headers=headers,params=params,verify=False)
    return response

def AllWorkstations():
    res = getWorkstations(100)
    res_dict = json.loads(res.text)
    workstations = []
    workstations += res_dict["workstations"]
    count = 100
    print(res_dict["list_info"])
    print(len(workstations))
    while res_dict["list_info"]["has_more_rows"]:
        res = getWorkstations(100, count)
        res_dict = json.loads(res.text)
        workstations += res_dict["workstations"]
        print(res_dict["list_info"])
        print(len(workstations))
        count += 100

    return workstations

ws = AllWorkstations()


def ppDict(input, depth = 0):
    tabs = " |" * depth
    if input == None:
        print(tabs + " None")
    elif type(input) is dict:
        for key in input:
            print(tabs + " " + key)
            ppDict(input[key], depth + 1)
    
    elif type(input) is list:
        for item in input:
            ppDict(item, depth + 1)

    else:
        print(tabs + " " + str(input))

def updateBitlocker(x, key):
    url = f"http://10.3.1.220:443/api/v3/workstations/{x}"
    headers = {"authtoken":"7C6702E0-321E-417D-BF35-8DD79B1BE8FD"}

    item_dict = {
        "workstation": {
            "workstation_udf_fields": {
            "udf_sline_601": key
            }
        }
    }

    item_json = json.dumps(item_dict)

    data = {'input_data': item_json}
    response = requests.put(url,headers=headers,data=data,verify=False)
    return response


errors = []
success = []
for item in ws:
    try:
        updateBitlocker(item["id"], data[item["computer_system"]["service_tag"]])
        print(f"Workstation {item['id']} {item['computer_system']['service_tag']} updated: {data[item['computer_system']['service_tag']]}")
        success.append(item['computer_system']['service_tag'])

    except:
        print(f"ERROR ADDING {item['computer_system']['service_tag']} KEY")
        errors.append(item['computer_system']['service_tag'])
