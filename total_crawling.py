from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.keys import Keys

from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time 
import re 

import requests
from urllib.request import Request, urlopen
from bs4 import BeautifulSoup
import pandas as pd 

'''
          필수 설치 정보
------------------------------------
최종 파일 저장 경로      : ".//results//"
chromedriver            : 자동업데이트 패키지 로드로 코드 수정
필수 설치 패키지         : pip install selenium, pip install BeautifulSoup, pip install pandas



           코드 정보
------------------------------------
순서대로 실행해야함, 2번 항목이 완료된 경우에는 3번만 진행해도 됨
# 수행완료 시간 : 최소 1일

# 1. 강좌의 대중소 분류 카테고리를 추출하는 코드
# (결과물) kocw_강의분류.xlsx
Lecture_categorys_df = Lecture_categorys()

# 2. 수집할 강좌의 기초정보를 추추하는 코드(각 강좌별로 부여된 세부링크를 수집하기 위한 코드)   
# > 소요시간이 오래걸림ㅜ 수집 중간마다 데이터 셋이 저장되도록 코드를 작성함
# (결과물) kocw_소분류단위_강의링크_정보.xlsx
Lecture_infos_df = Lecture_infos(Lecture_categorys_df)

# 3. 수집할 강좌의 기초정보를 통해 분석에서 정의한 필수 데이터들을 수집하는 코드 
# (결과물) kocw_강의정보_데이터셋.xlsx
# 수집된 파일로 재실행할때 예시코드
# Lecture_infos_df = pd.read_excel(".//results//kocw_소분류단위_강의링크_정보.xlsx")   > idx = idx+1 로 변경
Lecture_final_df = Lecture_final_Dataset(Lecture_infos_df)

'''




def get_category_id(idx):  
    category_id = idx + 1
    if category_id < 10 : category_ids = f'0{category_id}'
    else :  category_ids = category_id
    return str(category_id) , str(category_ids)
# 1. 강좌의 대중소 분류 카테고리를 추출하는 코드
def Lecture_categorys():
    url = "http://www.kocw.net/home/search/majorCourses.do"
    response = requests.get(url)

    # 결과물 표
    results = { "대분류" : "",
                "중분류" : "",
                "소분류" : "",
                "대분류 ID" : "",
                "중분류 ID" : "",
                "소분류 ID" : "",   
                "소분류 개수" : "",
                }

    # 저장 파일
    final_df = pd.DataFrame()

    def get_category_id(idx):  
        category_id = idx + 1
        if category_id < 10 : category_ids = f'0{category_id}'
        else :  category_ids = category_id
        return str(category_id) , str(category_ids)

    if response.status_code == 200:
        html = response.text
        soup = BeautifulSoup(html,"html.parser")
        
        # 카테고리 개수 구하기
        categorys = soup.select("ul.leftM > li ")
        
        for idm, category in enumerate(categorys):
            # 대분류 구하기
            Main_category_id = get_category_id(idm)[0]  # 1 
            Main_category_ids = get_category_id(idm)[1] # 01
            Main_category = soup.select(f'li#lev1Menu{str(Main_category_id)} > a')[0].get_text().split(" ")[0]        

            # 중분류 구하기
            Main_categorys =  soup.select(f"li#lev1Menu{str(Main_category_id)} > ul.leftS > li ")
            print(len(Main_categorys))
            # Middle_categorys = soup.select(f'ul.leftS > li#lev2Menu{str(Main_category_id)} > a')
            for idmm, Mcategory in enumerate(Main_categorys):
                Middle_category_ids = Main_category_ids + get_category_id(idmm)[1]
                Middle_categoryList = soup.select(f'li#lev1Menu{str(Main_category_id)} > ul.leftS > li#lev2Menu{str((idmm+1))} > a')#[0].get_text().split(" ")[0]
        
                for idx, Middle_category in enumerate(Middle_categoryList): 
                    Middle_category = Middle_category.get_text().split(" ")[0]      
                    print(Middle_category)
                
                    # 소분류 구하기
                    sub_categoryList  = soup.select(f'li#lev1Menu{str(Main_category_id)} > ul.leftS > li#lev2Menu{str((idmm+1))} > ul.leftL > li')
                    # print(sub_categoryList)
                    
                    for ids, sub_categorys in enumerate(sub_categoryList):  
                        sub_category_ids = sub_categorys.find("a")["id"]
                        sub_category = sub_categorys.get_text().split(" ")[0]
                        sub_category_count = re.sub(r"[^0-9]", "",  sub_categorys.get_text().split("(")[1])
                        

                        results = { "대분류" : Main_category,
                                    "중분류" : Middle_category,
                                    "소분류" : sub_category, 
                                    "대분류 ID" : Main_category_ids,
                                    "중분류 ID" : Middle_category_ids,
                                    "소분류 ID" : sub_category_ids,
                                    "소분류 개수" : sub_category_count,   
                                    }
                        
                        resultsDf= pd.DataFrame(results, index=[0])

                        final_df = pd.concat([final_df, resultsDf])


            
    final_df.reset_index(drop=True, inplace=True)
    final_df.index += 1
    final_df.index.name="IDX"

    final_df.to_excel(".//results//kocw_강의분류.xlsx", engine="xlsxwriter")
    
    return final_df
# 2. 수집할 강좌의 기초정보를 추추하는 코드(각 강좌별로 부여된 세부링크를 수집하기 위한 코드)
def Lecture_infos(Lecture_categorys_df):
    chrome_options = Options()
    # chrome_options.add_argument('--headless')
    chrome_options.add_argument('--no-sandbox')
    chrome_options.add_argument('--disable-dev-shm-usage')
    # 브라우저 윈도우 사이즈
    chrome_options.add_argument('window-size=1920x1080')
    # 드라이버 위치 경로 자동생성 패키지 입력
    # 참고 - https://yeko90.tistory.com/entry/%EC%85%80%EB%A0%88%EB%8B%88%EC%9B%80-%EA%B8%B0%EC%B4%88-executablepath-has-been-deprecated-please-pass-in-a-Service-object-%EC%97%90%EB%9F%AC-%ED%95%B4%EA%B2%B0-%EB%B0%A9%EB%B2%95
    driver = webdriver.Chrome(service = Service(ChromeDriverManager().install()))
    
    # driver.implicitly_wait(3)

    results = { # cattegory_crawling에서 수집
                "대분류"  : "",
                "중분류"  : "",
                "소분류"  : "",
                "대분류 ID" : "",
                "중분류 ID" : "",
                "소분류 ID" : "",   
                "소분류 개수" : "",
                
                # 1차 수집 진행
                "강좌링크" : "", # LectureLink
                "강좌명"  : "",  # LectureTitle
                
                # 2차 수집 진행
                "대학"  : "",
                "교수"  : "",
                "강의학기"  : "",
                "조회수"  : "",
                "차시별 강의"  : "",
                "강의설명"  : "",
                }


    # 소분류별 강의목표 정리
    df = pd.DataFrame(columns = list(results.keys())[:-6])

    # 소분류 ID를 통해 강의 정보 조회
    Lecture_categorys = Lecture_categorys_df
    sub_category_counts = Lecture_categorys["소분류 개수"]
    sub_category_IDs = Lecture_categorys["소분류 ID"]

    for idx, sub_category_ID in enumerate(sub_category_IDs) :
        idx = idx+1
        # 페이지네이션을 위한 설정
        sub_category_count = re.sub(r"[^0-9]", "", sub_category_counts[idx])
        pageNation = round(int(sub_category_count)/10)
        
        # 소분류와 관련된 기초 정보
        Lecture_category = Lecture_categorys.loc[idx]
        print(Lecture_category)
        
        for pageNumber in range(pageNation+1):
            print("첫번째")
            pageNumber = pageNumber + 1
            print(pageNumber)
            # base url 정보
            searchUrl = f'http://www.kocw.net/home/search/majorCourses.do#subject/{sub_category_ID}/pn/{pageNumber}'
            
            time.sleep(2.5)
            print("세번째")
            driver.get(searchUrl)
            html = driver.page_source
            soup = BeautifulSoup(html,"html.parser")
            # print(soup)
            
            # 소분류 강의 리스트  li:nth-child(2) > dl > dd > dl > dd:nth-child(4)
            LectureList = soup.select("dl.listCon2 > dt > strong > a")
            # LectureSummarys =  soup.select("ul.lectContList > li:nth-child(2) > dl > dd > dl > dd:nth-child(4)")
            
            for idxx , Lecture in enumerate(LectureList) :
                idxx = idxx+1
                LectureLink = Lecture.attrs['href']
                LectureTitle = Lecture.get_text()
                
                # summary_dd = soup.select(f"ul.lectContList > li:nth-child({idxx}) > dl > dd > dl > dd:nth-child(4)")
                
                # if len(summary_dd) != 0 : 
                #     LectureSummary = summary_dd[0].get_text()
                #     print(LectureSummary)
                # else : LectureSummary = ""
                #pageAj2 > div > div.searchResultWrap > 
                
                results = { "강좌명" : LectureTitle,
                            "강좌링크" : LectureLink,
                            # "강의설명" : LectureSummary,
                            }
                
                resultsDf= pd.DataFrame(results, index=[0])
                print(resultsDf)
                # 강좌 기초 정보 입력
                Lecture_category_df = pd.DataFrame(Lecture_category)
                Lecture_category_dict = Lecture_category_df.to_dict()
                # print(Lecture_category_df)
                
                # 3
                df = pd.concat([df,resultsDf]) 
                fill_values = Lecture_category_dict[list(Lecture_category_dict.keys())[0]]
                df = df.fillna(fill_values)

                print(df)

                df.drop_duplicates(inplace=True)
                df.reset_index(drop=True, inplace=True)
                df.index += 1
                df.index.name="IDX"
                df.to_excel(".//results//kocw_소분류단위_강의링크_정보.xlsx", engine="xlsxwriter")
                
    
    return df
# 3. 수집할 강좌의 기초정보를 통해 분석에서 정의한 필수 데이터들을 수집하는 코드 
def Lecture_final_Dataset(Lecture_infos_df):
    results = { # cattegory_crawling에서 수집
                "대분류"  : "",
                "중분류"  : "",
                "소분류"  : "",
                "대분류 ID" : "",
                "중분류 ID" : "",
                "소분류 ID" : "",   
                "소분류 개수" : "",
                
                # 1차 수집 진행
                "강좌링크" : "", # LectureLink
                "강좌명"  : "",  # LectureTitle
                
                # 2차 수집 진행
                "대학"  : "",
                "교수"  : "",
                "강의학기"  : "",
                "조회수"  : "",
                "차시별 강의"  : "",
                "차시별 개수"  : "",
                "강의설명"  : "",
                }


    # 소분류별 강의목표 정리
    df = pd.DataFrame(columns = list(results.keys()))


    # 소분류 링크를 통해 강의 정보 조회
    Lecture_infos = Lecture_infos_df

    for idx, LectureLink in enumerate(Lecture_infos["강좌링크"]):
        idx = idx+1
        # 소분류와 관련된 기초 정보
        Lecture_info = Lecture_infos.loc[idx]
        print(Lecture_info)
        
        # base url 정보
        searchUrl = f'http://www.kocw.net'+ LectureLink
        response = requests.get(searchUrl)

        if response.status_code == 200:
            html = response.text
            soup = BeautifulSoup(html,"html.parser")
            
            college = soup.select("ul.detailTitInfo > li:nth-child(1)")
            professor =  soup.select("ul.detailTitInfo > li:nth-child(2)")
            lecture_semester = soup.select("ul:nth-child(2) > li:nth-child(2) > dl > dd")
            
            if len(lecture_semester)== 0 :
                lecture_semester = soup.select("ul:nth-child(1) > li:nth-child(2) > dl > dd")
                
            views = soup.select("ul:nth-child(3) > li:nth-child(1) > dl > dd")
            if len(views)== 0 :
                views = soup.select("ul:nth-child(2) > li:nth-child(1) > dl > dd")
                
            summary_dd = soup.select(f"div.resultDetailWrap > div.detailViewStyle01 > div.datailViewInfo")
            
            if len(summary_dd) != 0 : 
                LectureSummary = summary_dd[0].get_text()
                print(LectureSummary)
            else : LectureSummary = ""

            
            lecture_class_table = soup.find('table',{"class":"tbType01"})
            
            if isinstance(lecture_class_table, type(None))!= True:
                lecture_class_nums = lecture_class_table.find_all("td",{"class":"no"})
                
                lecture_class_names = []
                
                for idx, lecture_class_num in enumerate(lecture_class_nums):
                    idx = idx+1
                    lecture_class_no = lecture_class_num.get_text()[:-1]
                    lecture_class_name = soup.select(f'#aTitle{lecture_class_no}')
                    if not lecture_class_name : pass
                    else : 
                        lecture_class_name = re.sub(r"[^\n\uAC00-\uD7A30-9a-zA-Z\s]", "",lecture_class_name[0].get_text()).strip()
                        lecture_class_names.append(lecture_class_no +". "+lecture_class_name)
                
            try  : 
                results = { "대학"  : college[0].get_text(),
                            "교수"  : professor[0].get_text(),
                            "강의학기"  : [lecture.get_text() if isinstance(lecture, type(None))!= True else 0 for lecture in lecture_semester ][0],
                            "조회수"  : [re.sub(r"[^0-9]", "", view.get_text()) if isinstance(view, type(None))!= True  else 0 for view in views][0],
                            "강의설명" : re.sub(r"[^\n\uAC00-\uD7A30-9a-zA-Z\s]", " ",LectureSummary),
                            "차시별 강의"  : str(lecture_class_names),
                            "차시별 개수"  : len(lecture_class_names),
                            }
                print(results)
                resultsDf= pd.DataFrame.from_dict([results])
                print(resultsDf)
                # 강좌 기초 정보 입력
                Lecture_info_df = pd.DataFrame(Lecture_info)
                Lecture_info_dict = Lecture_info_df.to_dict()
                # print(Lecture_category_df)
                
                # 3
                df = pd.concat([df,resultsDf]) 
                fill_values = Lecture_info_dict[list(Lecture_info_dict.keys())[0]]
                df = df.fillna(fill_values)
                
            except  : 
                print("sosgik96@naver.com으로 연락 ㄱㄱ")

    print(df)

    df.drop_duplicates(inplace=True)
    df.reset_index(drop=True, inplace=True)
    df.index += 1
    df.index.name="IDX"
    df.to_excel(".//results//kocw_강의정보_데이터셋.xlsx", engine="xlsxwriter")
    
    return df


# 1. 강좌의 대중소 분류 카테고리를 추출하는 코드
Lecture_categorys_df = Lecture_categorys()
# 2. 수집할 강좌의 기초정보를 추추하는 코드(각 강좌별로 부여된 세부링크를 수집하기 위한 코드)
Lecture_infos_df = Lecture_infos(Lecture_categorys_df)
# 3. 수집할 강좌의 기초정보를 통해 분석에서 정의한 필수 데이터들을 수집하는 코드 
# 수집된 파일로 재실행할때 예시코드
# Lecture_infos_df = pd.read_excel(".//results//kocw_소분류단위_강의링크_정보.xlsx")   > idx = idx+1 로 변경
Lecture_final_df = Lecture_final_Dataset(Lecture_infos_df)