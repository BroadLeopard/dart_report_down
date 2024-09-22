import os
import time
import tabula
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select

down_folder_link = None

def search_download(_company_name, _year, _data):
    global down_folder_link
    down_folder_link = str(_company_name) + '_' + str(_year) + '년_' + str(_data) + '_' + time.strftime('%Y_%m_%d_%H_%M_%S')

    # Chrome 드라이버 생성
    options = webdriver.ChromeOptions()
    # headless 옵션 설정
    # options.add_argument('headless')
    # 브라우저 윈도우 사이즈
    options.add_argument('window-size=1920x1080')

    #절대경로가 아닌 상대경로로 바꿀 수 있으면 더 좋음
    options.add_experimental_option("prefs", {"download.default_directory": os.path.dirname(os.path.abspath(__file__)) + '\\' + down_folder_link})

    driver = webdriver.Chrome(options=options)
    driver.get('https://dart.fss.or.kr/dsab007/main.do?option=corp')
    driver.implicitly_wait(5)


    #검색어 박스에 회사 이름 입력
    driver.find_element(By.XPATH, '//*[@id="textCrpNm"]').send_keys(_company_name)

    #검색 연도 설정 기본값 1년
    if year == 3:
        driver.find_element(By.XPATH, '// *[ @ id = "date5"]').click()
    elif year == 5:
        driver.find_element(By.XPATH, '// *[ @ id = "date6"]').click()
    elif year == 10:
        driver.find_element(By.XPATH, '// *[ @ id = "date7"]').click()

    #정기공시 박스 클릭
    driver.find_element(By.XPATH, '//*[@id="li_01"]').click()

    #사업보고서 체크박스 클릭
    driver.find_element(By.XPATH, '//*[@id="publicTypeDetail_A001"]').click()
    #분기보고서 체크박스 클릭
    driver.find_element(By.XPATH, '//*[@id="publicTypeDetail_A002"]').click()
    #반기보고서 체크박스 클릭
    driver.find_element(By.XPATH, '//*[@id="publicTypeDetail_A003"]').click()

    #조회 건수 최대 50
    Select(driver.find_element(By.XPATH, '//*[@id="maxResultsCb"]')).select_by_index(2)
    #검색 버튼 누르기
    driver.find_element(By.XPATH, '//*[@id="searchForm"]/div[1]/div/ul/li/div[3]/a').click()

    #로드에 많이 걸릴 수 있어서 5초 정도 대기
    time.sleep(5)


    #테이블 가져오기
    tbody = driver.find_element(By.XPATH, '//*[@id="tbody"]')

    if tbody:
        link_li = []
        down_link_li = []

        #받을 폴더 생성
        os.mkdir(down_folder_link)

        #2번째에 원하는 다운받고자 하는 링크 있음
        for tr in tbody.find_elements(By.TAG_NAME, 'tr'):
            #보고서 링크를 저장
            link_li.append(tr.find_elements(By.TAG_NAME, 'td')[2].find_element(By.TAG_NAME, 'a').get_attribute('href'))

        for link in link_li:
            driver.get(link)
            driver.implicitly_wait(5)

            #링크 쪼개서 인자만 획득
            tmp = driver.find_element(By.CLASS_NAME, 'btnDown').get_attribute('onclick').split('\'')

            #매우 위험한 방식
            rcpNo = tmp[1]
            dcmNo = tmp[3]

            #스크립트 함수에 있는 생성 방식 따라함
            down_link_li.append('https://dart.fss.or.kr/pdf/download/main.do?rcp_no='+rcpNo+'&dcm_no='+dcmNo)

        for down_link in down_link_li:
            driver.get(down_link)
            driver.implicitly_wait(5)
            
            #다운로드 버튼 누름
            driver.find_element(By.XPATH, '/html/body/div/div[3]/div/div/table/tbody/tr/td[2]/a').click()
            time.sleep(5)

        driver.quit()

def process_data(_data):
    os.chdir(down_folder_link)

    pdf_file_li = os.listdir(os.getcwd())

    #연도, 월
    tmp_li = []

    for pdf_file in pdf_file_li:
        #연도, 분기 추출
        tmp = pdf_file.split('.')
        tmp_li.append([int(tmp[0][-2:]), int(tmp[1][:2])])

        tmp_data_li = []

        tables = tabula.read_pdf(pdf_file, pages="all", stream=True)

        for index, table in enumerate(tables):
            print(f"\nData Index: {index}")
            print(type(table))
            print(table.head())

        #파일에서 검색어를 통해서 원하는 자료 추출. 형태는 [이름, 숫자]
        #한 번 정렬하고 추가
        #tmp_data_li.sort()
        #tmp_li[-1].extend(tmp_data_li)

    #tmp_li.sort() 정렬 기준은 연도, 분기
    #엑셀파일 생성

    for tmp_data in tmp_li:
        #여기에서 분기 한 번 Q 단위로 정리

        pass

if __name__ == "__main__":
    # 회사 이름 선택
    company_name = 'STX엔진'

    # 기간 설정 1년 3년 5년
    year = 1

    # 데이터 설정
    data = ' '

    search_download(company_name, year, data)

    process_data(data)
