import os
import time
import threading
import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.common.exceptions import NoSuchElementException


class DictionarySearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("네이버 영어사전 단어장 생성기")
        self.root.geometry("600x600")
        self.root.resizable(True, True)

        # 변수 설정
        self.is_running = False
        self.driver = None
        self.word_entries = []
        self.word_frames = []
        self.save_path = os.path.join(os.path.expanduser("~"), "Desktop", "naver_dic")

        # 옵션 변수
        self.collect_meaning = tk.BooleanVar(value=True)
        self.collect_pronunciation = tk.BooleanVar(value=True)
        self.collect_verb_tense = tk.BooleanVar(value=True)
        self.collect_main_content = tk.BooleanVar(value=True)
        self.collect_synonym = tk.BooleanVar(value=True)
        self.collect_antonym = tk.BooleanVar(value=True)

        # UI 생성
        self.create_widgets()

    def create_widgets(self):
        # 메인 프레임
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # 상단 프레임 (저장 경로)
        path_frame = ttk.Frame(main_frame)
        path_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Label(path_frame, text="저장 경로:").pack(side=tk.LEFT)
        self.path_entry = ttk.Entry(path_frame, width=30)  # 길이 줄임
        self.path_entry.pack(side=tk.LEFT, padx=5)
        self.path_entry.insert(0, self.save_path)

        browse_btn = ttk.Button(path_frame, text="찾아보기", command=self.browse_save_location)
        browse_btn.pack(side=tk.LEFT, padx=5)

        # 중앙 프레임 (검색어 및 옵션)
        center_frame = ttk.Frame(main_frame)
        center_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 검색어 프레임
        search_frame = ttk.Frame(center_frame)
        search_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        ttk.Label(search_frame, text="검색할 단어 목록").pack(anchor=tk.W)

        # 스크롤 가능한 검색어 프레임 (높이 제한)
        self.words_canvas = tk.Canvas(search_frame, height=150)  # 높이 제한 추가
        scrollbar = ttk.Scrollbar(search_frame, orient=tk.VERTICAL, command=self.words_canvas.yview)
        self.scrollable_frame = ttk.Frame(self.words_canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.words_canvas.configure(
                scrollregion=self.words_canvas.bbox("all")
            )
        )

        self.words_canvas.create_window((0, 0), window=self.scrollable_frame, anchor=tk.NW)
        self.words_canvas.configure(yscrollcommand=scrollbar.set)

        # pack 메서드에서 expand=False로 설정하여 자동 확장 방지
        self.words_canvas.pack(side=tk.LEFT, fill=tk.X, expand=False)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 단어 추가 버튼
        add_word_btn = ttk.Button(search_frame, text="단어 추가", command=self.add_word_entry)
        add_word_btn.pack(anchor=tk.W, pady=5)

        # 첫 번째 단어 입력 필드 추가
        self.add_word_entry()

        # 옵션 프레임
        options_frame = ttk.LabelFrame(center_frame, text="수집 옵션")
        options_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(10, 0))

        ttk.Checkbutton(options_frame, text="한글 뜻", variable=self.collect_meaning).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(options_frame, text="발음", variable=self.collect_pronunciation).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(options_frame, text="동사변화", variable=self.collect_verb_tense).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(options_frame, text="세부내용", variable=self.collect_main_content).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(options_frame, text="유의어", variable=self.collect_synonym).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(options_frame, text="반의어", variable=self.collect_antonym).pack(anchor=tk.W, pady=2)

        # 로그 프레임 (최소 높이 설정)
        log_frame = ttk.LabelFrame(main_frame, text="로그")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        log_frame.pack_propagate(False)  # 내부 위젯에 의한 크기 조정 방지
        log_frame.configure(height=200)  # 최소 높이 설정

        self.log_text = scrolledtext.ScrolledText(log_frame, height=10, wrap=tk.WORD)
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # 버튼 프레임
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))

        self.start_btn = ttk.Button(button_frame, text="검색 시작", command=self.start_search)
        self.start_btn.pack(side=tk.LEFT, padx=5)

        self.stop_btn = ttk.Button(button_frame, text="검색 중지", command=self.stop_search, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=5)

        self.open_file_btn = ttk.Button(button_frame, text="저장된 파일 위치 열기", command=self.open_file_location)
        self.open_file_btn.pack(side=tk.LEFT, padx=5)

        # 상태 표시줄
        self.status_var = tk.StringVar(value="준비됨")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def add_word_entry(self):
        frame = ttk.Frame(self.scrollable_frame)
        frame.pack(fill=tk.X, pady=2)

        entry = ttk.Entry(frame, width=40)  # 길이 늘림
        entry.pack(side=tk.LEFT, padx=(0, 5), fill=tk.X, expand=True)

        delete_btn = ttk.Button(frame, text="삭제", width=5,
                                command=lambda f=frame, e=entry: self.remove_word_entry(f, e))
        delete_btn.pack(side=tk.LEFT)

        self.word_entries.append(entry)
        self.word_frames.append(frame)

        # 스크롤 영역 업데이트
        self.scrollable_frame.update_idletasks()
        self.words_canvas.configure(scrollregion=self.words_canvas.bbox("all"))

    def remove_word_entry(self, frame, entry):
        if len(self.word_entries) > 1:  # 최소 하나의 입력 필드는 유지
            if entry in self.word_entries:
                self.word_entries.remove(entry)
            if frame in self.word_frames:
                self.word_frames.remove(frame)
            frame.destroy()

            # 스크롤 영역 업데이트
            self.scrollable_frame.update_idletasks()
            self.words_canvas.configure(scrollregion=self.words_canvas.bbox("all"))

    def browse_save_location(self):
        folder_path = filedialog.askdirectory(initialdir=self.path_entry.get())
        if folder_path:
            self.path_entry.delete(0, tk.END)
            self.path_entry.insert(0, folder_path)

    def log(self, message):
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)

    def start_search(self):
        # 입력된 단어 가져오기
        words = [entry.get().strip() for entry in self.word_entries if entry.get().strip()]

        if not words:
            messagebox.showwarning("경고", "검색할 단어를 입력해주세요.")
            return

        # 저장 경로 확인
        save_path = self.path_entry.get()
        if not os.path.exists(save_path):
            try:
                os.makedirs(save_path)
                self.log(f"저장 경로 생성: {save_path}")
            except Exception as e:
                messagebox.showerror("오류", f"저장 경로를 생성할 수 없습니다: {str(e)}")
                return

        # UI 상태 업데이트
        self.is_running = True
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.status_var.set("검색 중...")

        # 검색 스레드 시작
        threading.Thread(target=self.search_words, args=(words, save_path), daemon=True).start()

    def stop_search(self):
        self.is_running = False
        self.log("사용자에 의해 검색이 중지되었습니다.")
        self.status_var.set("중지됨")

        if self.driver:
            try:
                self.driver.quit()
            except:
                pass
            self.driver = None

        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

    def open_file_location(self):
        save_path = self.path_entry.get()
        excel_path = os.path.join(save_path, "사전검색결과.xlsx")

        if os.path.exists(save_path):
            try:
                if os.name == 'nt':  # Windows
                    os.startfile(save_path)
                elif os.name == 'posix':  # macOS, Linux
                    import subprocess
                    subprocess.Popen(['open', save_path])
            except Exception as e:
                self.log(f"파일 위치를 열 수 없습니다: {str(e)}")
        else:
            messagebox.showinfo("알림", "저장 폴더가 존재하지 않습니다.")

    def search_words(self, words, save_path):
        try:
            self.log(f"총 {len(words)}개의 단어에 대한 검색을 시작합니다.")

            # 데이터 저장 리스트 (직관적인 변수명으로 변경)
            english_words = []  # 영어단어
            korean_meanings = []  # 한글 뜻
            pronunciations = []  # 발음
            verb_changes = []  # 동사변화
            detailed_contents = []  # 세부내용
            synonyms = []  # 유의어
            antonyms = []  # 반의어
            related_words = []  # 파생어(관련어)

            # 웹드라이버 초기화
            options = Options()
            options.add_argument("--headless=new")
            self.driver = webdriver.Chrome(options=options)

            for i, keyword in enumerate(words):
                if not self.is_running:
                    break

                self.log(f"[{i + 1}/{len(words)}] '{keyword}' 검색 중...")
                self.status_var.set(f"검색 중... ({i + 1}/{len(words)})")

                # 기본 데이터 추가
                english_words.append(keyword)

                try:
                    # 네이버 사전 검색
                    url = f"https://en.dict.naver.com/#/search?query={keyword}"
                    self.driver.get(url)
                    time.sleep(2)

                    try:
                        self.driver.find_elements(By.CLASS_NAME, "highlight")[0].click()
                        time.sleep(3)
                    except:
                        self.log(f"'{keyword}' 검색 결과 클릭 실패")

                    # 1. 한글 뜻
                    if self.collect_meaning.get():
                        try:
                            mean = self.driver.find_element(By.CLASS_NAME, "entry_mean_list").text
                            if mean:
                                self.log(f"한글 뜻 수집 완료")
                                korean_meanings.append(mean)
                            else:
                                self.log(f"한글 뜻 검색 결과가 없습니다")
                                korean_meanings.append("-")  # 빈 문자열 대신 "-" 사용
                        except:
                            self.log(f"한글 뜻 검색 결과가 없습니다")
                            korean_meanings.append("-")  # 빈 문자열 대신 "-" 사용
                    else:
                        korean_meanings.append("-")  # 빈 문자열 대신 "-" 사용

                    # 발음
                    if self.collect_pronunciation.get():
                        try:
                            pronouns = self.driver.find_element(By.CLASS_NAME, "pronounce_area").text
                            pronouns = pronouns.replace('\n', " ")
                            if pronouns:
                                self.log(f"발음 수집 완료")
                                pronunciations.append(pronouns)
                            else:
                                self.log(f"발음 결과가 없습니다")
                                pronunciations.append("-")  # 빈 문자열 대신 "-" 사용
                        except:
                            self.log(f"발음 결과가 없습니다")
                            pronunciations.append("-")  # 빈 문자열 대신 "-" 사용
                    else:
                        pronunciations.append("-")  # 빈 문자열 대신 "-" 사용

                    # 2. 동사변화
                    if self.collect_verb_tense.get():
                        try:
                            try:
                                verb_tense = self.driver.find_element(By.XPATH,
                                                                      "//*[@id='content']/div[2]/div/div[4]/dl/dd/div").text
                            except:
                                verb_tense = self.driver.find_element(By.XPATH,
                                                                      "//*[@id='content']/div[2]/div/div[5]/dl/dd/div").text

                            if verb_tense:
                                verb_tense = verb_tense.replace('어휘등급', "")
                                verb_tense = verb_tense.replace('\n', " ")
                                self.log(f"동사변화 수집 완료")
                                verb_changes.append(verb_tense)
                            else:
                                self.log(f"동사변화 검색 결과가 없습니다")
                                verb_changes.append("-")  # 빈 문자열 대신 "-" 사용
                        except:
                            self.log(f"동사변화 검색 결과가 없습니다")
                            verb_changes.append("-")  # 빈 문자열 대신 "-" 사용
                    else:
                        verb_changes.append("-")  # 빈 문자열 대신 "-" 사용

                    # 3. 세부내용
                    if self.collect_main_content.get():
                        try:
                            main_con = self.driver.find_element(By.XPATH, "//*[@id='content']/div[4]/div[1]/div").text
                            if main_con:
                                try:
                                    main_con = main_con.replace("예문 열기", "")
                                    main_con = main_con.replace("학습 정보", "")
                                    main_con = main_con.replace("영영 사전", "")
                                    main_con = main_con.replace("영영사전", "")
                                    main_con = main_con.replace("\n", " ")
                                except:
                                    pass
                                self.log(f"세부내용 수집 완료")
                                detailed_contents.append(main_con)
                            else:
                                self.log(f"세부내용이 없습니다")
                                detailed_contents.append("-")  # 빈 문자열 대신 "-" 사용
                        except:
                            self.log(f"세부내용이 없습니다")
                            detailed_contents.append("-")  # 빈 문자열 대신 "-" 사용
                    else:
                        detailed_contents.append("-")  # 빈 문자열 대신 "-" 사용

                    # 4. 유의어
                    if self.collect_synonym.get():
                        try:
                            synonym = self.driver.find_element(By.XPATH, "//*[@id='searchPage_thesaurus']").text
                            if synonym:
                                try:
                                    synonym = synonym.replace("Oxford Thesaurus", "")
                                    synonym = synonym.replace("Collins Gem Thesaurus", "")
                                    synonym = synonym.replace(" 더보기  of English 더보기", "")
                                    synonym = synonym.replace("영영 사전", "")
                                    synonym = synonym.replace("of English 더보기", "")
                                    synonym = synonym.replace("\n", " ")
                                except:
                                    pass
                                self.log(f"유의어 수집 완료")
                                synonyms.append(synonym)
                            else:
                                self.log(f"유의어가 없습니다")
                                synonyms.append("-")  # 빈 문자열 대신 "-" 사용
                        except:
                            self.log(f"유의어가 없습니다")
                            synonyms.append("-")  # 빈 문자열 대신 "-" 사용
                    else:
                        synonyms.append("-")  # 빈 문자열 대신 "-" 사용

                    # 5. 반의어 (thesaurus.com에서 수집)
                    if self.collect_antonym.get():
                        try:
                            url = f"https://www.thesaurus.com/browse/{keyword}"
                            self.driver.get(url)
                            time.sleep(2)

                            # 새로운 HTML 구조에 맞게 반의어 추출
                            antonym_list = []
                            antonym_buttons = self.driver.find_elements(By.CSS_SELECTOR,
                                                                        "button.ZIu4Pmp_JmWEfr4WEkmw:not(.rXhgGek5ucMOQ0I0J3pa)")

                            for button in antonym_buttons:
                                if "Antonyms" in button.text:
                                    button.click()
                                    time.sleep(1)
                                    antonym_section = self.driver.find_element(By.CSS_SELECTOR,
                                                                               "div.QXhVD4zXdAnJKNytqXmK")
                                    antonym_links = antonym_section.find_elements(By.TAG_NAME, "a")

                                    for link in antonym_links:
                                        antonym_list.append(link.text)

                            if antonym_list:
                                antonym_text = ", ".join(antonym_list)
                                self.log(f"반의어 수집 완료")
                                antonyms.append(antonym_text)
                            else:
                                self.log(f"반의어가 없습니다")
                                antonyms.append("-")  # 빈 문자열 대신 "-" 사용
                        except Exception as e:
                            self.log(f"반의어 수집 중 오류: {str(e)}")
                            antonyms.append("-")  # 빈 문자열 대신 "-" 사용
                    else:
                        antonyms.append("-")  # 빈 문자열 대신 "-" 사용

                    # 6. 파생어 (제거된 기능)
                    related_words.append("-")  # 빈 문자열 대신 "-" 사용

                except Exception as e:
                    self.log(f"'{keyword}' 검색 중 오류 발생: {str(e)}")

                    # 누락된 데이터 채우기
                    if len(english_words) <= i:
                        english_words.append(keyword)
                    if len(korean_meanings) <= i:
                        korean_meanings.append("-")  # 빈 문자열 대신 "-" 사용
                    if len(pronunciations) <= i:
                        pronunciations.append("-")  # 빈 문자열 대신 "-" 사용
                    if len(verb_changes) <= i:
                        verb_changes.append("-")  # 빈 문자열 대신 "-" 사용
                    if len(detailed_contents) <= i:
                        detailed_contents.append("-")  # 빈 문자열 대신 "-" 사용
                    if len(synonyms) <= i:
                        synonyms.append("-")  # 빈 문자열 대신 "-" 사용
                    if len(antonyms) <= i:
                        antonyms.append("-")  # 빈 문자열 대신 "-" 사용
                    if len(related_words) <= i:
                        related_words.append("-")  # 빈 문자열 대신 "-" 사용

            # 웹드라이버 종료
            if self.driver:
                self.driver.quit()
                self.driver = None

            if not self.is_running:
                return

            # 데이터프레임 생성 및 저장
            self.log("검색 완료. 엑셀 파일 생성 중...")
            data_frame = pd.DataFrame([
                english_words,  # 영어단어
                korean_meanings,  # 한글 뜻
                pronunciations,  # 발음
                verb_changes,  # 동사변화
                detailed_contents,  # 세부내용
                synonyms,  # 유의어
                antonyms,  # 반의어
                related_words  # 파생어(관련어)
            ])
            data_frame_re = data_frame.transpose()
            data_frame_re.columns = ['영어단어', '한글 뜻', '발음', '동사변화', '세부내용', '유의어', '반의어', '파생어(관련어)']

            excel_path = os.path.join(save_path, "사전검색결과.xlsx")
            data_frame_re.to_excel(excel_path)

            self.log(f"작업이 완료되었습니다. 파일 저장 위치: {excel_path}")

            # UI 상태 업데이트
            self.root.after(0, self.search_completed)

        except Exception as e:
            self.log(f"검색 중 오류 발생: {str(e)}")
            if self.driver:
                self.driver.quit()
                self.driver = None

            # UI 상태 업데이트
            self.root.after(0, self.search_failed)

    def search_completed(self):
        self.is_running = False
        self.status_var.set("완료됨")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        messagebox.showinfo("완료", "단어 검색이 완료되었습니다.")

    def search_failed(self):
        self.is_running = False
        self.status_var.set("오류 발생")
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)


def main():
    root = tk.Tk()
    app = DictionarySearchApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
