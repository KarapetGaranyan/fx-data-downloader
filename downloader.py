import os
import win32com.client as client
import datetime, time
import requests


date = (datetime.datetime.today() - datetime.timedelta(days=2)).strftime("%Y%m%d")
input_path = os.getcwd() + "/old version/"
EUR, GBP, AUD, CAD, JPY, XAU, XAG = "58", "42", "37", "48", "69", "437", "458",
urls = [EUR, GBP, AUD, CAD, JPY, XAU, XAG]
curr = ["EUR", "GBP", "AUD", "CAD", "JPY", "XAU", "XAG"]

def get_files():
    if len(os.listdir(input_path)) == 0:
        for i, j in zip(urls, curr):
            session = requests.session()
            url = 'https://www.cmegroup.com/CmeWS/exp/voiProductDetailsViewExport.ctl'
            headers = {
            'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'accept-language': 'ru,en;q=0.9,fr;q=0.8',
            'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/98.0.4758.141 YaBrowser/22.3.3.889 Yowser/2.5 Safari/537.36'
            }
            payload = {'media':'xls',
                'tradeDate': date,
                'reportType':'P',
                'productId': i
            }
            resp = session.get(url, params=payload, headers=headers)
            with open(input_path + f'{j}.xls', 'wb') as wf:
                wf.write(resp.content)

def files_converter():
    if len(os.listdir(input_path)) > 0:
        excel = client.Dispatch("excel.application")
        for file in os.listdir(input_path):
            filename, fileextension = os.path.splitext(file)
            wb = excel.Workbooks.Open(input_path + file)
            output_path = os.getcwd() + "/new version/" + filename
            wb.SaveAs(output_path, 51)
            wb.Close
            remove_file = input_path + file
            os.remove(remove_file)
        excel.Quit()


# os.system('curl -A "Firefox/119" -H "Accept-Encoding: x" "https://www.cmegroup.com/CmeWS/exp/voiProductDetailsViewExport.ctl?media=xls&tradeDate=20240126&reportType=P&productId=58" -o VoiDetailsForProduct.xls')