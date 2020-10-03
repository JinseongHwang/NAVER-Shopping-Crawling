import sys
import os.path
import io
import time
import re
import math
from openpyxl import Workbook
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# 인코딩 타입 명시
# sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding='utf-8')
# sys.stderr = io.TextIOWrapper(sys.stderr.detach(), encoding='utf-8')

xlpath = 'C:\\Users\\User\\Desktop\\Python_Workspace\\devmoomin\\NAVER_SHOPPING.xlsx'
wb = Workbook()
ws = wb.active
ws.append(['ID', '상호명', '대표자', '고객센터', '사업자등록번호', '사업장 소재지', '통신판매업번호', 'e-mail', 'URL'])

# chrome이랑 chrome driver의 버전이 같아야한다.
driver = webdriver.Chrome('chromedriver')
driver.get('https://search.shopping.naver.com/mall/mall.nhn')

# 스마트 스토어 설정 체크
driver.find_element_by_id('gift_shopn').click()
time.sleep(1)

base_url = 'smartstore.naver.com/'

url_array = []

# 파싱할 페이지 범위 설정 (BEGIN ~ END - 1)
# MAX PAGE = 15000 (END MAX is 15001)
BEGIN, END = 1234, 1250

for page in range(BEGIN, END):
    print(f'\n>>> 현재 {page}page URL들을 파싱합니다.')

    # page 넘기기
    driver.execute_script(f'mall.changePage({page}, "listPaging");')
    time.sleep(1)

    # 몰 테이블 태그 접속
    mall_list = driver.find_element_by_class_name('malltv_lst')
    mall_array = mall_list.find_elements_by_css_selector('table > tbody > tr')

    for i in range(1, len(mall_array)):
        title = mall_array[i].find_elements_by_tag_name('a')[1].text
        try:
            curr_url = mall_array[i].find_element_by_partial_link_text(base_url).text
            store_id = curr_url[21:]
            seller_info_url = f'https://{base_url}{store_id}/profile?cp=2'
            url_array.append(seller_info_url)

        except Exception as e:
            print(f' ! [예외 발생] "{title}"은 "{base_url}" 형식이 아닙니다.')


print('\n>>> 사업자 정보 파싱을 시작합니다.')

# 순서대로 정수, 사업자등록번호, 통신판매업번호, 이메일 매칭 정규식
integer_regex = re.compile('^[0-9]+$')
businum_regex = re.compile('^[0-9]*-?[0-9]*$')
selling_code_regex = re.compile(r'[\s0-9제]{1, 7}-[\s0-9가-힣]{1, 7}-[\s0-9호]{1, 7}|[\s0-9가-힣]{1, 5}-?[\s0-9가-힣]{1, 10}|\(간이과세자 - 신고의무면제\)')
email_regex = re.compile(r'[a-zA-Z0-9+-_.]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+')

parse_cnt, exception_cnt = 0, 0

for url in url_array:
    # 진행상황 알림 + id counting
    parse_cnt += 1
    print(f'{parse_cnt}/{len(url_array)} # {url}')
    
    driver.get(url)

    # 모바일 페이지로 넘어가면 원하는 class name이 바뀌는 현상을
    # desktop 페이지로 변경해줘서 해결 (url에 m. 을 제거함)
    redirect_url = driver.current_url
    if redirect_url[8] == 'm':
        pre_url = redirect_url[:8]
        post_url = redirect_url[10:]
        desktop_url = pre_url + post_url
        driver.get(desktop_url)
        url = desktop_url

    shop_name = ceo = service_center_number = business_number = business_place = selling_code = email = None
    finished = False

    try:
        data_box = driver.find_element_by_class_name('oSdeQo13Wd')
        div_arr = data_box.find_elements_by_tag_name('div')

        # for i, v in enumerate(div_arr):
        #     print(i, v.text, sep=': ')
        # print(div_arr[18].text)
            
        shop_name = div_arr[4].text # 상호명
        ceo = div_arr[7].text # 대표자
        service_center_number = div_arr[11].text # 고객센터
        if integer_regex.match(service_center_number[-1]) is None:
            service_center_number = service_center_number[:-2]

        idx = 15 # div_arr의 index가 15인 값들에 대한 처리
        curr = div_arr[idx].text # 사업자 등록번호
        if businum_regex.match(curr):
            business_number = curr
        else:    
            business_place = curr # 사업장 소재지
            
        idx = 18 # div_arr의 index가 18인 값들에 대한 처리
        curr = div_arr[idx].text
        if email_regex.match(curr):
            email = curr
            finished = True # 항상 email이 마지막 요소이기 때문이다.
        elif selling_code_regex.match(curr):
            selling_code = curr
        else:
            business_place = curr
        
        if finished is False:
            idx = 21 # div_arr의 index가 21인 값들에 대한 처리
            curr = div_arr[idx].text
            if email_regex.match(curr):
                email = curr
                finished = True
            else:
                selling_code = curr
            
        if finished is False:
            idx = 24 # div_arr의 index가 24인 값들에 대한 처리
            email = div_arr[idx].text
            finished = True

        if shop_name in ceo: # 상호명을 대표자로 해놓은 경우
            shop_name = driver.find_element_by_class_name('_6P7lESLavN').text

        curr_business_info = [parse_cnt, shop_name, ceo, service_center_number, business_number, business_place, selling_code, email, url]
        ws.append(curr_business_info)
        # print(parse_cnt, shop_name, ceo, service_center_number, business_number, business_place, selling_code, email, url, sep='\n')
    
    except Exception as e:
        print(f' ! [예외 발생] "{url}" 에서 예외가 발생했습니다!')
        parse_cnt -= 1
        exception_cnt += 1


wb.save(xlpath)
driver.close()

print(f'>>> {exception_cnt}개의 예외를 제외한 {parse_cnt}개의 사업자 정보 파싱을 완료했습니다.')

sys.exit()