import requests
from bs4 import BeautifulSoup
import openpyxl
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from urllib.parse import urlparse, parse_qs, urlunparse, urlencode

def clean_query_text(query_text):
    """
    queryText에서 불필요한 패턴을 제거하는 함수
    """
    # 제거할 불필요한 패턴
    remove_patterns = ['znAll,', '@op,rsRESEARCH']
    
    # 각 패턴을 query_text에서 제거
    for pattern in remove_patterns:
        query_text = query_text.replace(pattern, '')
    return query_text

def get_colname_description(col_name):
    """
    colName 값에 따라 설명을 반환하는 함수
    """
    descriptions = {
        'bib_t': '학위논문',
        're_a_kor': '학술논문',
        're_a_over': '해외학술논문',
        'bib_m': '단행본',
        're_t': '연구보고서',
        'kem': '공개강의',
    }
    return descriptions.get(col_name, '알 수 없는 자료')

def extract_query_text(url):
    """
    URL에서 queryText 값을 추출하는 함수
    """
    from urllib.parse import urlparse, parse_qs

    parsed_url = urlparse(url)
    query_params = parse_qs(parsed_url.query)
    query_text = query_params.get('queryText', [''])[0]
    return query_text

def fetch_data(url, progress, progress_label):
    try:
        # 초기 설정
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Number", "Title", "Span1", "Span2", "Span3", "..."])
        
          # URL에서 colName과 queryText 값 추출
        from urllib.parse import urlparse, parse_qs
        parsed_url = urlparse(url)
        query_params = parse_qs(parsed_url.query)
        col_name = query_params.get('colName', [''])[0]
        col_description = get_colname_description(col_name)
        query_text = clean_query_text(extract_query_text(url))

        # 페이지 수 계산
        page = 0
        start_count = 0
        total_items = 0

        while True:
            response = requests.get(f"{url}&iStartCount={start_count}")
            soup = BeautifulSoup(response.text, 'html.parser')
            items = soup.select('.srchResultListW > ul > li')
            print(items)
            if not items:
                break
           
            for index, item in enumerate(items):
                title = item.select_one('.title').text.strip() if item.select_one('.title') else ''
                spans = item.select('.etc > span')
                span_texts = [span.text.strip() for span in spans]
                # sheet.append([start_count + index + 1, title] + span_texts)
                sheet.append([start_count + index + 1, title] + span_texts + [query_text, col_description])
               
                total_items += 1

            start_count += 10
            page += 1
            
            # 진행도 업데이트
            progress['value'] = page * 10  # 예시 진행도 계산
            progress_label.config(text=f"Progress: {total_items} items processed")
            root.update_idletasks()

        # 결과 저장
        workbook.save("output.xlsx")
        messagebox.showinfo("Completed", "Data has been saved to output.xlsx")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {str(e)}")


def modify_i_start_count(url):
    i_start_count = "iStartCount=0"
    # URL에 'iStartCount=0'이 있으면 제거
    if i_start_count in url:
        url = url.replace(i_start_count, "")
        print(f"URL after removing 'iStartCount=0': {url}")

    return url

def modify_page_scale(url):
    # 찾고자 하는 문자열
    old_page_scale = "pageScale=10"
    new_page_scale = "pageScale=100"

    # URL에 'pageScale=10'이 있으면 'pageScale=100'으로 변경
    if old_page_scale in url:
        modified_url = url.replace(old_page_scale, new_page_scale)
        return modified_url
    else:
        messagebox.showwarning("Warning", "The URL does not contain 'pageScale=10'.")
        return url
    
def validate_and_modify_url(url):
    # 1. /search/Search.do? 의 코드가 있는지 확인
    search_path = '/search/Search.do?'
    if search_path in url:
        # 2. /search/Search.do? 있을 경우 앞에 url 정보 다 자르고 앞에 https://www.riss.kr/ 를 붙이기
        modified_url = 'https://www.riss.kr/search/Search.do?' + url.split(search_path, 1)[1]
        return modified_url
    else:
        # URL 형식이 올바르지 않을 경우 경고 메시지
        messagebox.showwarning("Invalid URL", "The URL does not contain '/search/Search.do?'. Please enter a valid RISS URL.")
        return None
    
def start_process():
    url = url_entry.get()
    if not url:
        messagebox.showwarning("Warning", "Please enter a valid URL.")
        return
    
    # URL 확인 및 수정
    modified_url = validate_and_modify_url(url)
    if not modified_url:
        return
    modified_url = modify_i_start_count(modified_url)
    print(f"Validated URL: {modified_url}")
    progress['value'] = 0
    progress_label.config(text="Starting...")
    fetch_data(modified_url, progress, progress_label)



# GUI 설정
root = tk.Tk()
root.title("RISS Data Fetcher")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

# URL 입력
url_label = tk.Label(frame, text="Enter RISS URL:")
url_label.grid(row=0, column=0, pady=5)
url_entry = tk.Entry(frame, width=50)
url_entry.grid(row=0, column=1, pady=5)

# 실행 버튼
start_button = tk.Button(frame, text="Start", command=start_process)
start_button.grid(row=1, column=1, pady=10)

# 진행도 표시
progress = ttk.Progressbar(frame, length=400, mode='determinate')
progress.grid(row=2, column=1, pady=10)
progress_label = tk.Label(frame, text="Progress: 0%")
progress_label.grid(row=3, column=1, pady=5)

# 프로그램 종료 버튼
exit_button = tk.Button(frame, text="Exit", command=root.quit)
exit_button.grid(row=4, column=1, pady=10)

root.mainloop()
