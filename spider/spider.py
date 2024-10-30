import requests
import json
import gzip
import datetime

url = "https://pfsc.agri.cn/api/priceQuotationController/pageList?key=&order="

headers = {
    "accept": "application/json, text/plain, */*",
    "accept-language": "en-US,en;q=0.8",
    "content-type": "application/json;charset=UTF-8",
    "sec-ch-ua": '"Chromium";v="130", "Brave";v="130", "Not?A_Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
    "sec-fetch-dest": "empty",
    "sec-fetch-mode": "cors",
    "sec-fetch-site": "same-origin",
    "sec-gpc": "1",
    "Referer": "https://pfsc.agri.cn/"
}

# First request to get total number of records
payload = {
    "pageNum": 1,
    "pageSize": 1,
    "marketId": "",
    "provinceCode": "",
    "pid": "",
    "varietyId": ""
}

response = requests.post(url, headers=headers, json=payload, verify=False)
total_records = response.json()['content']['total']

# Second request to get all records
payload['pageSize'] = total_records

response = requests.post(url, headers=headers, json=payload, verify=False)
data_list = response.json()['content']['list']

# Generate a better file name with timestamp
filename = datetime.datetime.now().strftime("data_%Y%m%d_%H%M%S.json.gz")

# Save the list directly to a compressed file
with gzip.open(filename, 'wt', encoding='utf-8') as f_out:
    json.dump(data_list, f_out, ensure_ascii=False, indent=4)