import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

def collect_books(base_url, delay=0.5):
    """
    Собирает данные о книгах с сайта books.toscrape.com и сохраняет в Excel.

    Args:
        base_url (str): URL первой страницы сайта (например, 'http://books.toscrape.com').
        delay (float): Задержка между запросами.

    Returns:
        None. Сохраняет данные в файл Excel.
    """

    page_num = 1
    data = []

    while True:
        if page_num == 1:
            url = f"{base_url}"  # Первая страница
        else:
            url = f"{base_url}catalogue/page-{page_num}.html"

        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
            'Accept-Language': 'en-US,en;q=0.9',
            'Referer': 'http://books.toscrape.com/',
            'DNT': '1',
            'Upgrade-Insecure-Requests': '1'
        }

        try:
            response = requests.get(url, headers=headers)
            response.raise_for_status()
        except requests.exceptions.RequestException as e:
            print(f"Ошибка при запросе {url}: {e}")
            if response.status_code == 404: #Проверяем, что ошибка именно 404
               print("Достигнут конец каталога. Завершаем сбор данных.")
               break #Прекращаем цикл, если 404 ошибка
            else:
                break #Прерываем цикл при любой другой ошибке

        soup = BeautifulSoup(response.text, 'lxml')
        articles = soup.find_all('article', class_='product_pod')

        if not articles:
            print("Больше нет книг на этой странице. Завершаем сбор данных.")
            break

        for article in articles:
            try:
                title = article.h3.a['title']
                price = article.find('p', class_='price_color').text.strip('£')
                rating_element = article.find('p', class_='star-rating')
                rating = rating_element['class'][1] if rating_element else 'No rating'

                data.append({'title': title, 'price': price, 'rating': rating})

            except Exception as e:
                print(f"Ошибка при обработке книги: {e}")

        page_num += 1
        time.sleep(delay)

    df = pd.DataFrame(data)
    df.to_excel('books_toscrape.xlsx', index=False)  # Сохраняем в Excel
    print("Данные успешно собраны и сохранены в books_toscrape.xlsx")


base_url = 'http://books.toscrape.com/'
collect_books(base_url)