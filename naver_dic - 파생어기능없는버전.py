from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd

d_00 = []
d_0 = []
d_1 = []
d_2 = []
d_3 = []
d_4 = []
d_5 = []
d_6 = []

def main():
    url = "https://en.dict.naver.com/#/search?query=" + keyword
    driver.get(url)
    time.sleep(2)
    d_0.append(keyword)

    try:
        driver.find_elements(By.CLASS_NAME, "highlight")[0].click()
    except:
        print("실패")
    time.sleep(3)

    # 1
    try:
        mean = driver.find_element(By.CLASS_NAME, "entry_mean_list").text
        if len(mean) > 1:
            print("")
            print(str(mean))
            d_1.append(mean)
        else:
            print("")
            print("한글 뜻 검색 결과가 없습니다")
            d_1.append("")
    except:
        print("")
        print("한글 뜻 검색 결과가 없습니다")
        d_1.append("")

    # 00
    pronouns = driver.find_element(By.CLASS_NAME, "pronounce_area").text
    pronouns = pronouns.replace('\n', " ")
    if len(pronouns) > 1:
        print("")
        print(str(pronouns))
        d_00.append(pronouns)
    else:
        print("")
        print("발음 결과가 없습니다")
        d_00.append("")

    # 2
    try:
        try:
            verb_tense = driver.find_element(By.XPATH, "//*[@id='content']/div[2]/div/div[4]/dl/dd/div").text
        except:
            verb_tense = driver.find_element(By.XPATH, "//*[@id='content']/div[2]/div/div[5]/dl/dd/div").text
        if len(verb_tense) > 1:
            print("")
            verb_tense = verb_tense.replace('어휘등급', "")
            verb_tense = verb_tense.replace('\n', " ")
            print(str(verb_tense))
            d_2.append(verb_tense)
        else:
            print("")
            print("동사변화 검색 결과가 없습니다")
            d_2.append("")
    except:
        print("")
        print("동사변화 검색 결과가 없습니다")
        d_2.append("")
    # 3
    main_con = driver.find_element(By.XPATH, "//*[@id='content']/div[4]/div[1]/div").text
    try:
        if len(main_con) > 1:
            try:
                main_con = main_con.replace("예문 열기", "")
                main_con = main_con.replace("학습 정보", "")
                main_con = main_con.replace("영영 사전", "")
                main_con = main_con.replace("영영사전", "")
                main_con = main_con.replace("\n", " ")
            except:
                pass
            print("")
            print(str(main_con))
            d_3.append(main_con)
        else:
            print("")
            print("메인 컨텐츠가 없습니다")
            d_3.append("")
    except:
        print("")
        print("메인 컨텐츠가 없습니다")
        d_3.append("")

    # 4 유의어
    try:
        synonym = driver.find_element(By.XPATH, "//*[@id='searchPage_thesaurus']").text
        if len(synonym) > 1:
            try:
                synonym = synonym.replace("Oxford Thesaurus", "")
                synonym = synonym.replace("Collins Gem Thesaurus", "")
                synonym = synonym.replace(" 더보기  of English 더보기", "")
                synonym = synonym.replace("영영 사전", "")
                synonym = synonym.replace("영영 사전", "")
                synonym = synonym.replace("of English 더보기", "")
                synonym = synonym.replace("\n", " ")
            except:
                pass
            print("")
            print(str(synonym))
            d_4.append(synonym)
        else:
            print("")
            print("유의어가 없습니다")
            d_4.append("")
    except:
        print("")
        print("유의어가 없습니다")
        d_4.append("")

    # 5 반의어
    url = "https://www.thesaurus.com/browse/" + keyword
    driver.get(url)
    time.sleep(2)
    try:
        antonym = driver.find_element(By.XPATH, "//*[@id='antonyms']/div[2]").text
        antonym = antonym.replace("\n", " ")
        if len(antonym) > 1:
            print("")
            print(str(antonym))
            d_5.append(antonym)
        else:
            print("")
            print("반의어가 없습니다")
            d_5.append("")
    except:
        print("")
        print("반의어가 없습니다")
        d_5.append("")

    # 6 파생어
    time.sleep(3)
    print("파생어 기능이 제거된 버전입니다.")
    d_6.append("")


if __name__ == '__main__':
    List = [x for x in input('\n검색할 단어 입력. 띄어쓰기로 구분합니다: \n').split()]
    print(str(len(List)) + '개의 항목에 대하여 검색을 시작합니다 \n')
    driver = webdriver.Chrome()
    for x in List:
        if str(x.isalpha()) == "False":
            print("저자명이 잘못되었습니다. 계속 진행시 프로그램이 강제종료 될 수 있습니다. 프로그램을 다시 시작해주세요")
            a = input()
        else:
            pass
        keyword = x
        main()
    data_frame = (pd.DataFrame([d_0, d_1, d_00, d_2, d_3, d_4, d_5, d_6]))
    data_frame_re = data_frame.transpose()
    data_frame_re.columns = ['영어단어', '한글 뜻', '발음', '동사변화', '세부내용', '유의어', '반의어', '파생어(관련어)']
    data_frame_re.to_excel("C:\\naver_dic\\사전검색결과.xlsx")

    print("")
    print('작업이 완료되었습니다')
