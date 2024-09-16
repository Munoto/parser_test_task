import time
import requests
from bs4 import BeautifulSoup
import xlsxwriter


def write_to_excel(data_list, filename='output.xlsx'):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()

    headers = ["Наименование на рус. языке", "БИН участника", "ФИО", "ИИН",
               "Полный адрес(рус)"]

    for col_num, header in enumerate(headers):
        worksheet.write(0, col_num, header)

    for row_num, data in enumerate(data_list, start=1):
        for col_num, header in enumerate(headers):
            worksheet.write(row_num, col_num, data.get(header, ""))

    workbook.close()


def parse_and_collect_data():
    url = "https://www.goszakup.gov.kz/ru/registry/rqc?count_record=2000&page=1"
    response = requests.get(url)

    soup = BeautifulSoup(response.text, "lxml")

    rows = soup.find_all('tr')

    links = []
    for row in rows:
        link_tags = row.find_all('a', href=True)
        for link in link_tags:
            href = link['href']
            full_url = requests.compat.urljoin(url, href)
            links.append(full_url)

    all_data = []
    count = 1
    for link in links:
        response = requests.get(link)
        page_soup = BeautifulSoup(response.text, "lxml")

        data = {
            "Наименование на рус. языке": "",
            "БИН участника": "",
            "ФИО": "",
            "ИИН": "",
            "Полный адрес(рус)": ""
        }

        rows = page_soup.find_all('tr')

        for row in rows:
            th = row.find('th')
            td = row.find('td')

            if th and td:
                field_name = th.text.strip()
                field_value = td.text.strip()

                if field_name in data and not data[field_name]:
                    data[field_name] = field_value

            cells = row.find_all('td')
            if len(cells) >= 3:
                full_address = cells[2].text.strip()

                if not data["Полный адрес(рус)"]:
                    data["Полный адрес(рус)"] = full_address

        if all(data.values()):
            all_data.append(data)
            print(count)
            count += 1

        time.sleep(1)


    return all_data


collected_data = parse_and_collect_data()
write_to_excel(collected_data, 'data.xlsx')
