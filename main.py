from pyquery import PyQuery as pq
import requests
import xlwings as xw


class Node(object):
    def __init__(self, id: str = None, bank_name: str = None, type: str = None):
        self.id = id
        self.bank_name = bank_name
        self.type = type


items = []
banks = []
posts = []


def init():
    print("--- 處理中 ---")
    app = xw.App(visible=True, add_book=False)
    wb = app.books.add()
    sheet = wb.sheets[0]
    sheet.title = "ogc"
    sheet.range('A1:A3').value = ["代號", "金融機構", "類型"]
    r = requests.get('http://www.esunbank.com.tw/event/announce/BankCode.htm')
    r.encoding = 'big5'
    doc = pq(r.text)
    style_name = {'style2': "銀行", "style5": "郵局", "style6": "信用合作社", "style3": "漁會", "style4": "農會"}
    credit_unions = []
    fishery_associations = []
    farmers = []
    trs = doc("table").eq(1).find("tr")
    sheet_index = 2
    for idx, tr in enumerate(trs):
        if idx == 0:
            continue
        tds = pq(tr).find("td")
        tds_index = len(tds)
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
            elif 'style3' == class_name:
                fishery_associations.append(Node(id_dom.text(), bank_dom.text(), style_name[class_name]))
            elif 'style4' == class_name:
                farmers.append(Node(id_dom.text(), bank_dom.text(), style_name[class_name]))
    total_list = banks + posts + credit_unions + fishery_associations + farmers
    for item in total_list:
        sheet.range('A' + str(sheet_index)).value = [item.id, item.bank_name, item.type]
        sheet_index += 1
    sheet.autofit()
    wb.save("banks.xlsx")
    app.quit()
    print("--- 處理結束 ---")


if __name__ == '__main__':
    init()
