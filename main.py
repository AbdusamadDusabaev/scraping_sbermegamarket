import time
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.wait import WebDriverWait
from selenium.webdriver.support import expected_conditions
from connect import get_unfinished_products, set_finished_product, record_product_data, set_unfinished_all_products
from connect import record_characteristics
import os
import shutil


ua_chrome = " ".join(["Mozilla/5.0 (Windows NT 10.0; Win64; x64)",
                      "AppleWebKit/537.36 (KHTML, like Gecko)",
                      "Chrome/108.0.0.0 Safari/537.36"])
options = webdriver.ChromeOptions()
options.add_argument(f"user-agent={ua_chrome}")
options.add_argument("--headless")
timeout = 30


def get_characteristics(bs_object):
    result = list()
    characteristics = bs_object.find_all(name="div", class_="pdp-specs__item")
    for characteristic in characteristics:
        name = characteristic.find(name="div", class_="pdp-specs__item-name").text.strip()
        value = characteristic.find(name="div", class_="pdp-specs__item-value")
        if value is not None:
            value = value.text.strip()
        else:
            value = characteristic.a.text.strip()
        sub_result = {"name": name, "value": value}
        result.append(sub_result)
    return result


def create_result_dir():
    if os.path.isdir("result"):
        shutil.rmtree("result")
        os.mkdir("result")
    else:
        os.mkdir("result")
    print("[INFO] Папка result успешно создана")


def get_info_about_product(browser, wait_driver, product):
    product_id = product["product_id"]
    print(f"[INFO] Запущен анализ товара с id = {product_id}")
    url = f"https://sbermegamarket.ru/catalog/?q={product_id}"
    browser.get(url=url)
    wait_driver.until(expected_conditions.visibility_of_element_located((By.CLASS_NAME, "breadcrumbs__content")))
    wait_driver.until(expected_conditions.visibility_of_element_located((By.TAG_NAME, "h1")))
    response = browser.page_source
    bs_object = BeautifulSoup(response, "lxml")
    title = bs_object.find(name="h1").text.strip()
    description = bs_object.find(name="div", class_="text-block", itemprop="description")
    if description is None:
        description = ""
    else:
        description = description.text.strip()
    product_link = browser.current_url
    product_code = product_link.split("/")[5].split("-")[-1]
    price = bs_object.find(name="meta", itemprop="price")["content"]
    currency = bs_object.find(name="meta", itemprop="priceCurrency")["content"]
    categories = bs_object.find(name="ul", class_="breadcrumbs__content")
    categories = categories.find_all(name="li", itemprop="itemListElement")
    category_one = categories[1].a.span.text.strip()
    category_two = categories[2].a.span.text.strip()
    category_three = categories[3].a.span.text.strip()
    category_four = categories[4].a.span.text.strip()
    file_name = " - ".join([category_one, category_two, category_three]) + ".xlsx"
    images = bs_object.find(name="div", class_="scroller__content scroller_enlarged")
    if images is not None:
        images = images.find_all(name="img")
        images = ", ".join([image["src"] for image in images])
    else:
        images = bs_object.find(name="div", class_="inner-image-zoom slide slide_extra-large").img["src"]
    index = product["index"]
    current_index = record_product_data(file_name=file_name, product_id=product_id, title=title,
                                        description=description, product_link=product_link, product_code=product_code,
                                        price=price, currency=currency, category_one=category_one,
                                        category_two=category_two, category_three=category_three,
                                        category_four=category_four, images=images)
    characteristics = get_characteristics(bs_object=bs_object)
    record_characteristics(file_name=file_name, index=current_index, characteristics=characteristics)
    set_finished_product(index=index)
    print(f"[INFO] Анализ товара с id = {product_id} закончен")


def parsing():
    browser = webdriver.Chrome(options=options)
    browser.set_window_size(width=1920, height=1080)
    wait_driver = WebDriverWait(driver=browser, timeout=timeout)
    try:
        products = get_unfinished_products()
        for product in products:
            if product["product_id"] is not None:
                start_time = time.time()
                try:
                    get_info_about_product(browser=browser, product=product, wait_driver=wait_driver)
                except Exception as ex:
                    print("[INFO] Не удалось получить данные о товаре. Продолжаем парсинг...")
                stop_time = time.time()
                print(f"[INFO] На обработку товара ушло {stop_time - start_time} секунд")
    finally:
        browser.close()
        browser.quit()


def main():
    mode = input("[INPUT] Продолжить парсинг или начать сначала? (1 - сначала, 2 - продолжить): >>> ")
    if mode == "1":
        set_unfinished_all_products()
        create_result_dir()
    parsing()
    parsing()
    parsing()
    print("[INFO] Парсинг товаров закончен")


if __name__ == "__main__":
    main()
