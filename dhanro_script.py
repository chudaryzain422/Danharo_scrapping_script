from selenium import webdriver
from bs4 import BeautifulSoup
import time
from urllib.parse import urljoin

import pandas as pd

from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import os

# Create an instance of Options
options = Options()

# Add any desired options
options.add_argument("--ignore-local-proxy")

# Create the webdriver instance
driver = webdriver.Chrome(options=options)

# Login details
login_url = 'https://danharoo.com/member/login.html'
username = 'gyu2016'
password = '!rlawjddjs123'

file_name = 'product_details.xlsx'
if not os.path.isfile(file_name):
    df = pd.DataFrame(columns=[
        'No',
        '판매 진행 여부',
        '등록제외 사유',
        '작성자',
        '원본 상품명',
        '상품명_필수',
        '상품명_100자',
        '검색 상품명',
        'Item number code',
        'Single item code',
        '모델NO',
        '상품약어',
        '상품약어_사방넷',
        '모델명_사방넷',
        '모델명',
        '모델명2',
        '모델NO_사방넷',
        '브랜드명_필수',
        '브랜드명_사방넷',
        '자체상품코드',
        '사이트검색어',
        '상품구분_필수',
        '카테고리',  # need to add here
        '매입처ID',  # same here
        '물류처ID_사방넷',
        '제조사_필수',
        '제조사_사방넷',
        '원산지_제조국_필수',
        '원산지_제조국_사방넷',
        '생산연도',
        '제조일',
        '시즌',
        '남녀구분',
        '상품상태_필수',
        '판매지역',
        '세금구분_필수',
        '배송비구분_필수',
        '배송비',
        '공급처지불배송비',
        '합배송가능_개수',
        '개당안전_마진',
        '배송비_사방넷',
        '반품지구분',
        'message1',
        'message2',
        'message3',
        'message4',
        'quantity',  # need to add here
        'period',
        'percent',
        'price1',
        '원가_필수',
        '판매가_필수',
        '원가_사방넷',
        '판매가_사방넷',
        '판매_수수료',
        '상품_마진',
        '상품_마진율',
        '배송비_마진',
        '총_마진_1개_판매시',
        '마진율_1개_판매시',
        '권장_판매가_계산식',
        '마진_계산식_배송비_적용',
        '마진율_계산식_배송비_적용',
        '최소_마진율_설정_기준금액에_대한_마진_시트에_설정된_마진율로_설정된다',
        '최소_마진율_권장_판매가',
        '최소_마진율_마진',
        '합배송_가능_개수에_의한_최소_마진',
        '최종권장_판매가',
        '최종_마진',
        '최종마진율',
        'TAG가_필수',
        '옵션제목_1',
        '옵션상세명칭_1_사방넷',
    ])
    df.to_excel(file_name, index=False)


try:
    # Open the login page
    driver.get(login_url)

    # Allow the page to load
    time.sleep(2)

    # Enter login details and submit
    driver.find_element(By.NAME, 'member_id').send_keys(username)
    driver.find_element(By.NAME, 'member_passwd').send_keys(password)
    driver.find_element(By.XPATH, '//a[@class="loginBtn -mov"]').click()

    # Allow the login process to complete
    time.sleep(2)

    category_urls = [
        'https://danharoo.com/product/list.html?cate_no=45',  # Category 1 URL
        'https://danharoo.com/product/list.html?cate_no=46',  # Category 2 URL
        'https://danharoo.com/product/list.html?cate_no=47',  # Category 2 URL
        'https://danharoo.com/product/list.html?cate_no=48',  # Category 2 URL
        'https://danharoo.com/product/list.html?cate_no=56',  # Category 2 URL
        'https://danharoo.com/product/list.html?cate_no=176',  # Category 2 URL
        'https://danharoo.com/product/list.html?cate_no=69',  # Category 2 URL
        'https://danharoo.com/product/list.html?cate_no=49',  # Category 2 URL
        'https://danharoo.com/product/list.html?cate_no=50',  # Category 2 URL
        # Add more category URLs here as needed
    ]

    # Now we are logged in, we can proceed to scrape the product pages
    for category_url in category_urls:
        driver.get(category_url)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        pagination = soup.find('div', class_='xans-element- xans-product xans-product-normalpaging ec-base-paginate')
        page_links = pagination.find_all('a', class_='other')
        page_numbers = [int(link.text.strip()) for link in page_links]
        max_page_number = max(page_numbers)

        for page_number in range(1, max_page_number + 1):
            url = f'{category_url}&page={page_number}'
            try:
                driver.get(url)
            except NoSuchElementException:
                print("Element not found. Exiting the loop.")
                break

            try:
                WebDriverWait(driver, 10).until(EC.url_to_be(url))
            except TimeoutException:
                print("Timeout exception occurred. Moving to the next page.")
                continue

            # Allow the page to load
            time.sleep(5)

            # Now, parse the page with BeautifulSoup
            soup = BeautifulSoup(driver.page_source, 'html.parser')
            grid_container = soup.find('ul', class_='prdList grid4')
            products = grid_container.find_all('li', class_='item xans-record-')
            for product_item in products:
                try:
                    # Click on the product item
                    relative_link = product_item.find('a', href=True)['href']
                    product_link = urljoin(category_url, relative_link)

                    # Open the product link
                    driver.get(product_link)

                    # Allow the page to load
                    time.sleep(5)

                    # Parse the product detail page
                    soup = BeautifulSoup(driver.page_source, 'html.parser')
                    detail_area = soup.find('div', class_='detailArea')

                    # Find the infoArea element within detail_area
                    info_area = detail_area.find('div', class_='infoArea')
                    product_name = info_area.find('h2', class_='item_name').text.strip()

                    # Extract product price
                    product_price = info_area.find('tr', class_='price xans-record-').find(
                        'td').text.strip() if info_area.find('tr', class_='price xans-record-') else ''

                    consumer_price = info_area.find_all('tr', class_='xans-record-')[2].find('td').text.strip() if len(
                        info_area.find_all('tr', class_='xans-record-')) > 2 else ''

                    brief_desc = info_area.find_all('tr', class_='xans-record-')[3].find('td').text.strip() if len(
                        info_area.find_all('tr', class_='xans-record-')) > 3 else ''

                    summary = info_area.find_all('tr', class_='xans-record-')[4].find('td').text.strip() if len(
                        info_area.find_all('tr', class_='xans-record-')) > 4 else ''

                    product_code = info_area.find_all('tr', class_='xans-record-')[5].find('td').text.strip() if len(
                        info_area.find_all('tr', class_='xans-record-')) > 5 else ''

                    category = \
                    soup.find('div', class_='xans-element- xans-product xans-product-headcategory path').find_all('a')[
                        1].text if soup.find('div',
                                             class_='xans-element- xans-product xans-product-headcategory path') else ''

                    quantity = soup.find('p', class_='info').text if soup.find('p', class_='info') else ''

                    color_select = info_area.find('select', id='product_option_id1')
                    color_options = [option.text.strip() for option in
                                     color_select.find_all('option')[2:]] if color_select else []

                    # Extract size options
                    size_select = info_area.find('select', id='product_option_id2')
                    size_options = [option.text.strip() for option in
                                    size_select.find_all('option')[2:]] if size_select else []

                    # Extract image URLs
                    image_urls = []
                    image_div = detail_area.find('div',
                                                 class_='xans-element- xans-product xans-product-addimage listImg')
                    if image_div:
                        images = image_div.find_all('img')
                        for img in images:
                            src = img.get('src')
                            image_urls.append(src)

                    # Append the product details to the list
                    data = {
                        'No': '',
                        '판매 진행 여부': '',
                        '등록제외 사유': '',
                        '작성자': '',
                        '원본 상품명': product_name,
                        '상품명[필수]': product_name,
                        '상품명[100자]': '',
                        '검색 상품명': '',
                        'Item number code': '',
                        'Single item code': '',
                        '모델NO': '',
                        '상품약어': '',
                        '상품약어[사방넷]': '',
                        '모델명[사방넷]': '',
                        '모델명': product_code,
                        '모델명2': '',
                        '모델NO[사방넷]': '',
                        '브랜드명[필수]': '',
                        '브랜드명[사방넷]': '',
                        '자체상품코드': product_code,
                        '사이트검색어': '',
                        '상품구분[필수]': '',
                        '카테고리': category,  # need to add here
                        '매입처ID': category,  # same here
                        '물류처ID[사방넷]': '',
                        '제조사[필수]': '',
                        '제조사[사방넷]': '',
                        '원산지(제조국)[필수]': '',
                        '원산지(제조국)[사방넷]': '',
                        '생산연도': '',
                        '제조일': '',
                        '시즌': '',
                        '남녀구분': '',
                        '상품상태[필수]': '',
                        '판매지역': '',
                        '세금구분[필수]': '',
                        '배송비구분[필수]': '',
                        '배송비': '',
                        '공급처지불배송비': '',
                        '합배송가능 개수': '',
                        '개당안전 마진': '',
                        '배송비[사방넷]': '',
                        '반품지구분': '',
                        'message1': brief_desc,
                        'message2': summary,
                        'message3': '',
                        'message4': '',
                        'quantity': quantity,  # need to add here
                        'period': '',
                        'percent': '',
                        'price1': product_price,
                        '원가[필수]': consumer_price,
                        '판매가[필수]': product_price,
                        '원가[사방넷]': '',
                        '판매가[사방넷]': '',
                        '판매 수수료': '6%',
                        '상품 마진': '-',
                        '상품 마진율': '',
                        '배송비 마진': '-',
                        '총 마진(1개 판매시)': '-',
                        '마진율(1개 판매시)': '',
                        '권장 판매가(계산식)': '',
                        '마진(계산식)(배송비 적용)': '',
                        '마진율(계산식)(배송비 적용)': '',
                        '최소 마진율 설정(기준금액에 대한 마진 시트에 설정된 마진율로 설정된다)': '',
                        '최소 마진율권장 판매가': '',
                        '최소 마진율마진': '',
                        '합배송 가능 개수에 의한최소 마진': '',
                        '최종권장 판매가': '',
                        '최종 마진': '',
                        '최종마진율': '',
                        'TAG가[필수]': '',
                        '옵션제목(1)': '',
                        '옵션상세명칭(1)[사방넷]': '',
                    }

                    # Add color options to the dictionary
                    for i, (color_option, size_option) in enumerate(zip(color_options, size_options), start=1):
                        option = f"{color_option}_{size_option}"
                        data[f'옵션상세명칭 {i}'] = option

                    # Add image URLs to the dictionary
                    for i, image_url in enumerate(image_urls, start=1):
                        modified_url = image_url[2:]  # Remove the first two slashes
                        data[f'대표이미지[필수] {i}'] = modified_url


                    # Append the product details to the DataFrame
                    df = pd.DataFrame([data])
                    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
                        df.to_excel(writer, index=False, startrow=writer.sheets['Sheet1'].max_row,
                                    header=not writer.sheets)
                    with pd.ExcelWriter(file_name, engine='openpyxl', mode='a') as writer:
                        df.to_excel(writer, index=False, header=False, startrow=writer.sheets['Sheet1'].max_row)

                    # Go back to the product list page
                    driver.back()

                    # Wait for the grid container to be present again
                    WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.CLASS_NAME, 'prdList.grid4')))
                    time.sleep(2)

                except (NoSuchElementException, TimeoutException) as e:
                    print("An exception occurred:", str(e))
                    continue

except Exception as e:
    print("An exception occurred:", str(e))

finally:
    # Close the browser when done
    driver.quit()
