from pyquery import PyQuery as pq
import requests
from openpyxl import Workbook


class Node(object):
    def __init__(self, id: str = None, bank_name: str = None, type: str = None):
        self.id = id
        self.bank_name = bank_name
        self.type = type


items: list = []
banks: list = []
posts: list = []
url: str = 'http://www.esunbank.com.tw/event/announce/BankCode.htm'


def init() -> None:
    wb: Workbook = Workbook()
    ws = wb.active
    ws.title = "banks"
    ws.sheet_properties.tabColor = '00FF0000'
    print("--- 處理中 ---")
    ws.append(["代號", "金融機構", "類型"])
    r = requests.get(url)
    r.encoding = r.apparent_encoding
    doc = pq(r.text)
    style_name = {'style2': "銀行", "style5": "郵局", "style6": "信用合作社", "style3": "漁會", "style4": "農會"}
    credit_unions: list = []
    fishery_associations: list = []
    farmers: list = []
    trs = doc("table").eq(1).find("tr")
    sheet_index = 2
    for idx, tr in enumerate(trs):
        if idx == 0:
            continue
        tds = pq(tr).find("td")
        tds_index: int = len(tds)
        for td_idx in range(0, tds_index, 2):
            id_dom = pq(tds[td_idx])

            bank_dom = pq(tds[td_idx + 1])
            class_name = bank_dom.attr('class')

            if 'style2' == class_name:
                banks.append(Node(id_dom.text(), bank_dom.text(), style_name[class_name]))
            elif 'style5' == class_name:
                posts.append(Node(id_dom.text(), bank_dom.text(), style_name[class_name]))
            elif 'style6' == class_name:
                credit_unions.append(Node(id_dom.text(), bank_dom.text(), style_name[class_name]))
                credit_unions.append(Node(id_dom.text(), bank_dom.text(), style_name[class_name]))
            elif 'style3' == class_name:
                fishery_associations.append(Node(id_dom.text(), bank_dom.text(), style_name[class_name]))
            elif 'style4' == class_name:
                farmers.append(Node(id_dom.text(), bank_dom.text(), style_name[class_name]))
    total_list = banks + posts + credit_unions + fishery_associations + farmers
    for item in total_list:
        ws.append([int(item.id), item.bank_name, item.type])
        ws['A' + str(sheet_index)].number_format = '00#'
        sheet_index += 1

    wb.save("banks.xlsx")
    wb.close()
    print("--- 處理結束 ---")


if __name__ == '__main__':
    init()
