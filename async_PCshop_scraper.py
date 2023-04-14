import time

import requests
import xlsxwriter
from bs4 import BeautifulSoup
from fake_useragent import UserAgent
from aiocfscrape import CloudflareScraper
import asyncio

ua = UserAgent(browsers=["chrome", "opera", "firefox"])
URL_main = "https://pcshop.ua/noutbuki-i-aksessuari/notebooks"

laptops_data = []


def get_soup(url, params=None):
    headers = {"User-Agent": ua.random}

    r = requests.get(url, headers=headers, params=params)
    return BeautifulSoup(r.content, "lxml")


async def get_async_info(card_url):
    processor, ram = "", ""

    async with CloudflareScraper() as session:
        async with session.get(card_url) as response:
            card_soup = BeautifulSoup(await response.text(), "lxml")  # Soup of each cards of laptops

            try:
                image_link = card_soup.find("img", class_="product-main-carousel__image").get("src")

                table = card_soup.find("table", class_="properties").find_all("td")
                for i in range(0, len(table)):
                    if table[i].text == "Модель процесора":
                        processor = table[i + 1].text
                    if table[i].text == "Об'єм оперативної пам'яті":
                        ram = table[i + 1].text

            except Exception:
                image_link = "NONE"

            return {
                "image": image_link,
                "price": card_soup.find("span", class_="product-info__price").text,
                "producer": table[1].text,
                "diagonal": table[9].text.split(" ")[0],
                "processor": processor,
                "ram": ram
            }


async def create_tasks(soup):
    laptops = soup.find_all("a", class_="product-thumb")  # Cards of laptops on each page

    tasks = []  # List of functions
    for laptop in laptops:
        card_link = laptop.get("href")  # Getting a link to the card of laptop
        task = asyncio.create_task(get_async_info(card_link))
        tasks.append(task)

    return await asyncio.gather(*tasks)  # At each iteration we have 24 get_async_info functions, which run asynchronously


def write_to_file(page_xlsx, row, info, column=0):
    for laptop_info in info:
        page_xlsx.write(row, column, laptop_info["producer"])
        page_xlsx.write(row, column + 1, laptop_info["price"])
        page_xlsx.write(row, column + 2, laptop_info["diagonal"])
        page_xlsx.write(row, column + 3, laptop_info["processor"])
        page_xlsx.write(row, column + 4, laptop_info["ram"])
        page_xlsx.write(row, column + 5, laptop_info["image"])
        row += 1


def main():
    cur_time = time.perf_counter()

    #            === Creating xlsx-book ===
    book = xlsxwriter.Workbook("laptops_PC_shop.xlsx")  # Creation of xlsx book
    page_xlsx = book.add_worksheet("laptops")
    page_xlsx.set_column(0, 5, 15)

    # Format settings
    sheet_format = book.add_format()
    sheet_format.set_bold()

    page_xlsx.write(0, 0, "Производитель", sheet_format), page_xlsx.write(0, 1, "Цена", sheet_format)
    page_xlsx.write(0, 2, "Диагональ", sheet_format), page_xlsx.write(0, 3, "Процессор", sheet_format)
    page_xlsx.write(0, 4, "RAM", sheet_format), page_xlsx.write(0, 5, "Фото", sheet_format)

    row = 1

    #         === Main code ===
    main_soup = get_soup(URL_main)
    pages_num = int(main_soup.find_all("a", class_="pagi__link")[-1].get_text())

    for page in range(1, pages_num + 1):
        pages_soup = get_soup(URL_main, params={"page": page})  # Getting a soup of each page with 24 laptops
        print(pages_soup.find("title").text.strip())

        info = asyncio.run(create_tasks(pages_soup))  # At each iteration we get a dict with data from func "get_async_data"

        write_to_file(page_xlsx, row, info)
        row += 24

    book.close()

    end = time.perf_counter()
    print(f"It took {end - cur_time} seconds")


if __name__ == '__main__':
    main()
