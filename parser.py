import requests
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows


def get_categories():
    headers = {
        'sec-ch-ua': '"Not/A)Brand";v="8", "Chromium";v="126", "Google Chrome";v="126"',
        'Referer': 'https://www.wildberries.ru/',
        'DNT': '1',
        'sec-ch-ua-mobile': '?0',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36',
        'sec-ch-ua-platform': '"Windows"',
    }

    url = 'https://static-basket-01.wbbasket.ru/vol0/data/main-menu-by-ru-v2.json'
    try:
        response = requests.get(url, headers=headers)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.HTTPError as http_err:
        print(f"HTTP error occurred: {http_err}")
    except requests.exceptions.ConnectionError as conn_err:
        print(f"Connection error occurred: {conn_err}")
    except requests.exceptions.Timeout as timeout_err:
        print(f"Timeout error occurred: {timeout_err}")
    except requests.exceptions.RequestException as req_err:
        print(f"An error occurred: {req_err}")
    except ValueError as json_err:
        print(f"JSON decoding failed: {json_err}")
    return None


def parse_categories(categories, parent_id=None, level=0):
    items = []

    for category in categories:
        category_id = category['id']
        items.append({
            'id': category_id,
            'name': ' ' * (level * 4) + category['name'],
            'level': level,
            'parent_id': parent_id
        })

        if 'childs' in category:
            items.extend(parse_categories(category['childs'], parent_id=category_id, level=level + 1))

    return items


def main():
    wb = Workbook()
    wb.remove(wb.active)

    categories = get_categories()

    for category in categories:
        main_category_name = category['name']
        items = parse_categories([category])
        df = pd.DataFrame(items)
        ws = wb.create_sheet(title=main_category_name)

        for r in dataframe_to_rows(df, index=False, header=True):
            ws.append(r)

    wb.save("wb-categories.xlsx")


if __name__ == '__main__':
    main()
