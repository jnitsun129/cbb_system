from selenium import webdriver
import pandas as pd
from bs4 import BeautifulSoup
from scripts.utils import get_chromedriver_path


class Driver:
    def __init__(self):
        self.driver = webdriver.Chrome(
            get_chromedriver_path())

    def get(self, url: str):
        self.driver.get(url)
        self.driver.implicitly_wait(10)
        return self.driver.page_source

    def quit(self):
        self.driver.quit()


def fetch_kenpom(driver: Driver) -> pd.DataFrame:
    url = 'https://kenpom.com'
    try:
        html_content = driver.get(url)
        soup = BeautifulSoup(html_content, 'html.parser')
        table = soup.find('table', {'id': 'ratings-table'})
        headers = [header.get_text(strip=True)
                   for header in table.select('tr.thead2 th')]
        rows = []
        for tbody in table.find_all('tbody'):
            for row in tbody.find_all('tr'):
                cells = [cell.get_text(strip=True) for cell in row.find_all(
                    'td', class_=lambda x: x != 'td-right')]
                rows.append(cells)
        headers = headers[:len(rows[0])]
        df = pd.DataFrame(rows, columns=headers)
        with open('kp_names.csv', 'w') as file:
            df.to_csv(file)
    except:
        print('Connection failed')
    return df
