import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By


def main():
    driver.get(f"https://en.dict.naver.com/#/search?query={keyword}")
    time.sleep(2)
    d_0.append(keyword)

    try:
        driver.find_element(By.CLASS_NAME, "highlight").click()
        time.sleep(3)
    except:
        print("실패")

    get_element_text(By.CLASS_NAME, "entry_mean_list", d_1)
    get_element_text(By.CLASS_NAME, "pronounce_area", d_00, replace_newline=True)
    get_verb_tense()
    get_main_content()
    get_synonym()
    get_antonym()
    print("총" + str(len(List)) + "개 중에서 이전에 비해 작업 1개가 더 완료되었습니다. ")
    d_6.append("")


def get_element_text(by_type, element_name, target_list, replace_newline=False):
    try:
        text = driver.find_element(by_type, element_name).text
        if replace_newline:
            text = text.replace('\n', ' ')
        if text:
            target_list.append(text)
        else:
            target_list.append("")
    except:
        target_list.append("결과없음")


def get_verb_tense():
    try:
        verb_tense = driver.find_element(By.XPATH, "//*[@id='content']/div[2]/div/div[4]/dl/dd/div").text
    except:
        try:
            verb_tense = driver.find_element(By.XPATH, "//*[@id='content']/div[2]/div/div[5]/dl/dd/div").text
        except:
            verb_tense = "결과없음"
    verb_tense = verb_tense.replace('어휘등급', '').replace('\n', ' ')
    d_2.append(verb_tense if verb_tense else '')


def get_main_content():
    try:
        main_con = driver.find_element(By.XPATH, "//*[@id='content']/div[4]/div[1]/div").text
        main_con = main_con.replace("예문 열기", "").replace("학습 정보", "").replace(
        "영영 사전", "").replace("영영사전", "").replace("\n", " ")
    except:
        main_con = "결과없음"
    d_3.append(main_con if main_con else '')


def get_synonym():
    try:
        synonym = driver.find_element(By.XPATH, "//*[@id='searchPage_thesaurus']").text
        synonym = synonym.replace("Oxford Thesaurus", "").replace("Collins Gem Thesaurus", "").replace(
        " 더보기  of English 더보기", "").replace("영영 사전", "").replace("영영 사전", "").replace(
        "of English 더보기", "").replace("\n", " ")
    except:
        synonym = "결과없음"
    d_4.append(synonym if synonym else '')


def get_antonym():
    driver.get(f"https://www.thesaurus.com/browse/{keyword}")
    time.sleep(2)
    try:
        antonym = driver.find_element(By.XPATH, "//*[@id='antonyms']/div[2]").text
        antonym = antonym.replace("\n", " ")
    except:
        antonym = "결과없음"
    d_5.append(antonym if antonym else '')


if __name__ == '__main__':
    List = input('\n검색할 단어 입력. 띄어쓰기로 구분합니다: \n').split()
    print(f'{len(List)}개의 항목에 대하여 검색을 시작합니다 \n')
    driver = webdriver.Chrome()

    d_0, d_1, d_00, d_2, d_3, d_4, d_5, d_6 = [], [], [], [], [], [], [], []

    for x in List:
        if str(x.isalpha()) == "False":
            print("검색어가 잘못되었습니다. 프로그램을 다시 시작해주세요")
            break
        keyword = x
        main()

    df = pd.DataFrame([d_0, d_1, d_00, d_2, d_3, d_4, d_5, d_6]).T
    df.columns = ['영어단어', '한글 뜻', '발음', '동사변화', '세부내용', '유의어', '반의어', '파생어(관련어)']
    df.to_excel("사전검색결과.xlsx", index=False)

    print('작업이 완료되었습니다')
