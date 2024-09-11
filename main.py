import requests
from bs4 import BeautifulSoup
import openpyxl
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

def fetch_data(url, progress, progress_label):
    try:
        # 초기 설정
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.append(["Number", "Title", "Span1", "Span2", "Span3", "..."])
        
        # 페이지 수 계산
        page = 0
        start_count = 0
        total_items = 0

        while True:
            response = requests.get(f"{url}&iStartCount={start_count}")
            soup = BeautifulSoup(response.text, 'html.parser')
            items = soup.select('.srchResultListW > ul > li')
            if not items:
                break

            for index, item in enumerate(items):
                title = item.select_one('.title').text.strip() if item.select_one('.title') else ''
                spans = item.select('.etc > span')
                span_texts = [span.text.strip() for span in spans]
                sheet.append([start_count + index + 1, title] + span_texts)
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

def start_process():
    url = url_entry.get()
    if not url:
        messagebox.showwarning("Warning", "Please enter a valid URL.")
        return
    
    progress['value'] = 0
    progress_label.config(text="Starting...")
    fetch_data(url, progress, progress_label)

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
