
# =============================================
# 📌 임포트체커 (v250808 + 분석기능 + 중복표시 기능)
# 작성자: 서기대
# 주요 기능:
#   - 엑셀 기반 수용가 단말기 등록용 SQL VALUES 자동 생성
#   - 수용가 통계 분석 (수량, 항목 분류 등)
#   - 수용가번호 중복 항목 적색 표시 후 저장
# =============================================

import time
import os
import requests
import gspread
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime
from tkinter import ttk
from openpyxl import load_workbook
from openpyxl.styles import Font

# 📌 엑셀 컬럼명 정의 (20개 항목)
COLUMNS = [
    '수용가명', '수용가번호', '구주소', '신주소', '경도', '위도', '업종', '소속',
    '블록', '수용가 전화번호', '수용가 대상 년도', '검침원', '검침일', '계량기번호',
    '구경', '통신', '단말 부번호', '단말 주번호', '단말 회사', '단말 설치일'
]

SITE_SQ = 4
COMPANY_SQ = 9

def generate_sql_from_excel():
    file_path = filedialog.askopenfilename(title="엑셀 파일 선택", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        messagebox.showwarning("파일 선택", "파일이 선택되지 않았습니다.")
        return

    idx = site_combobox.current()
    if idx < 0:
        messagebox.showwarning("계정명 선택", "먼저 계정명을 선택하세요.")
        return
    try:
        selected_num_len = int(filtered_df.iloc[idx]['수용가번호길이'])
    except Exception:
        messagebox.showerror("계정명 오류", "선택한 계정의 수용가번호길이 값이 올바르지 않습니다.")
        return

    try:
        df = pd.read_excel(file_path, dtype=str).fillna('')
        df = df.iloc[:, 1:1+len(COLUMNS)]
        df.columns = COLUMNS

        admin_no_counts = df['수용가번호'].value_counts()
        duplicated_admin_nos = set(admin_no_counts[admin_no_counts > 1].index)

        values_list = []
        success_count = 0
        fail_count = 0
        success_admin_no_list = []

        validation_cases = [
            ("수용가번호 길이 검사", lambda row: len(row['수용가번호']) == selected_num_len, "수용가번호 길이 불일치"),
            ("수용가번호 중복 검사", lambda row: row['수용가번호'] not in duplicated_admin_nos, "수용가번호 중복"),
            ("수용가 전화번호 길이 검사", lambda row: len(row['수용가 전화번호']) < 14, "수용가 전화번호 13자리 초과"),
        ]

        for idx_row, row in df.iterrows():
            try:
                for case_num, (desc, check_func, err_msg) in enumerate(validation_cases, 1):
                    if not check_func(row):
                        raise ValueError(f"[CASE #{case_num}] {err_msg} (값: {row['수용가번호']})")

                def safe_float(val, col):
                    if val == '':
                        raise ValueError(f"{col} 값이 비어있음")
                    return float(val)

                def safe_int(val, col):
                    if val == '':
                        raise ValueError(f"{col} 값이 비어있음")
                    return int(val)

                values = (
                    row['수용가명'], row['수용가번호'], row['구주소'], row['신주소'],
                    safe_float(row['경도'], '경도'), safe_float(row['위도'], '위도'),
                    row['업종'], SITE_SQ, COMPANY_SQ,
                    row['수용가 전화번호'], row['수용가 대상 년도'], row['검침원'],
                    safe_int(row['검침일'], '검침일'), row['계량기번호'],
                    safe_int(row['구경'], '구경'), row['통신'], row['단말 부번호'],
                    row['단말 주번호'], COMPANY_SQ,
                    f"{pd.to_datetime(row['단말 설치일']).date()}"
                )

                formatted = "('{}', '{}', '{}', '{}', {}, {}, '{}', {}, {}, '{}', '{}', '{}', {}, '{}', {}, '{}', '{}', '{}', {}, '{}'::timestamp)".format(*values)
                values_list.append(formatted)
                success_admin_no_list.append(row['수용가번호'])
                success_count += 1

            except Exception as e:
                values_list.append(f"-- [ERROR #{idx_row+1}] {e}")
                fail_count += 1

        result_text.config(state=tk.NORMAL)
        result_text.delete(1.0, tk.END)
        result_text.insert(tk.END, "-- ✅ 임포트전 조회할 수용가목록")
        result_text.insert(tk.END, ",".join(f"'{x}'" for x in success_admin_no_list) + "")
        result_text.insert(tk.END, f"-- 총 {len(df)}개 중 {success_count}개 성공, {fail_count}개 실패")
        result_text.insert(tk.END, ",".join(values_list))
        result_text.config(state=tk.DISABLED)

    except Exception as e:
        messagebox.showerror("에러 발생", str(e))

def analyze_excel_customer_stats():
    file_path = filedialog.askopenfilename(title="엑셀 파일 선택", filetypes=[("Excel files", "*.xlsx *.xls")])
    if not file_path:
        messagebox.showwarning("파일 선택", "파일이 선택되지 않았습니다.")
        return

    try:
        df = pd.read_excel(file_path, engine='openpyxl').fillna('')
        df.columns = [col.strip() for col in df.columns]

        result_text.config(state=tk.NORMAL)
        result_text.delete(1.0, tk.END)

        # 실제 컬럼명 출력
        result_text.insert(tk.END, "📋 엑셀 파일의 실제 컬럼명:\n")
        for i, col in enumerate(df.columns, 1):
            result_text.insert(tk.END, f"{i:2d}. {col}\n")
        result_text.insert(tk.END, "\n")

        # 수용가번호 총 수량
        if '수용가번호' in df.columns:
            total = df['수용가번호'].nunique()
            result_text.insert(tk.END, f"✅ 수용가번호 총 수량: {total}\n\n")
        else:
            result_text.insert(tk.END, "⚠️ '수용가번호' 컬럼이 없습니다.\n\n")

        # 일반적인 분석 컬럼들 (실제 컬럼명에 맞게 조정)
        analysis_columns = []
        
        # 블록 관련 컬럼 찾기
        block_columns = [col for col in df.columns if '블록' in col]
        analysis_columns.extend(block_columns)
        
        # 구분 관련 컬럼 찾기
        division_columns = [col for col in df.columns if '구분' in col or '분류' in col or '구' in col]
        analysis_columns.extend(division_columns)
        
        # 기타 분석 가능한 컬럼들
        other_columns = ['업종', '소속', '통신', '구경', '검침원']
        for col in other_columns:
            if col in df.columns:
                analysis_columns.append(col)
        
        # 중복 제거
        analysis_columns = list(set(analysis_columns))
        
        if analysis_columns:
            result_text.insert(tk.END, "📊 항목별 통계 분석:\n")
            for col in analysis_columns:
                if col in df.columns:
                    result_text.insert(tk.END, f"\n📈 '{col}' 항목별 개수:\n")
                    counts = df[col].value_counts()
                    total_count = len(counts)
                    result_text.insert(tk.END, f"총 {total_count}개 항목\n")
                    
                    # 상위 10개만 표시 (너무 많으면 화면이 복잡해짐)
                    display_counts = counts.head(10)
                    for value, count in display_counts.items():
                        percentage = (count / len(df)) * 100
                        result_text.insert(tk.END, f"- {value}: {count}개 ({percentage:.1f}%)\n")
                    
                    if len(counts) > 10:
                        result_text.insert(tk.END, f"... 외 {len(counts) - 10}개 항목\n")
                else:
                    result_text.insert(tk.END, f"⚠️ 컬럼 '{col}' 없음\n")
        else:
            result_text.insert(tk.END, "⚠️ 분석 가능한 컬럼을 찾을 수 없습니다.\n")

        result_text.config(state=tk.DISABLED)

    except Exception as e:
        messagebox.showerror("분석 오류", str(e))

def mark_duplicates_in_place():
    file_path = filedialog.askopenfilename(title="엑셀 파일 선택", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        messagebox.showwarning("파일 선택", "파일이 선택되지 않았습니다.")
        return

    try:
        df = pd.read_excel(file_path, dtype=str).fillna('')
        if '수용가번호' not in df.columns:
            messagebox.showerror("열 없음", "'수용가번호' 열이 존재하지 않습니다.")
            return

        duplicated = df['수용가번호'][df['수용가번호'].duplicated(keep=False)]
        if duplicated.empty:
            messagebox.showinfo("중복 없음", "중복된 수용가번호가 없습니다.")
            return

        # 진행상태 창 생성
        progress_window = tk.Toplevel()
        progress_window.title("중복 표시 진행상태")
        progress_window.geometry("400x150")
        progress_window.transient(window)
        progress_window.grab_set()

        # 프로그레스바
        progress_label = tk.Label(progress_window, text="중복 수용가번호 적색 표시 중...")
        progress_label.pack(pady=(20, 10))
        
        progress_bar = ttk.Progressbar(progress_window, length=300, mode='determinate')
        progress_bar.pack(pady=(0, 10))
        
        status_label = tk.Label(progress_window, text="")
        status_label.pack(pady=(0, 10))

        wb = load_workbook(file_path)
        ws = wb.active
        headers = [cell.value for cell in ws[1]]
        col_idx = headers.index('수용가번호') + 1  # 1-based
        red_font = Font(color="FF0000")

        total_rows = ws.max_row - 1  # 헤더 제외
        progress_bar['maximum'] = total_rows

        for row in range(2, ws.max_row + 1):
            cell = ws.cell(row=row, column=col_idx)
            if str(cell.value) in duplicated.values:
                cell.font = red_font
            
            # 진행상태 업데이트
            progress_bar['value'] = row - 1
            status_label.config(text=f"처리 중: {row-1}/{total_rows} 행")
            progress_window.update()

        wb.save(file_path)
        progress_window.destroy()
        
        # 필터링 옵션 제공
        response = messagebox.askyesno("필터링 옵션", 
                                     f"중복 수용가번호가 적색으로 표시되어 저장되었습니다.\n\n"
                                     f"중복된 수용가번호만 포함된 새 파일을 생성하시겠습니까?")
        
        if response:
            create_filtered_file(file_path, duplicated.values)

    except Exception as e:
        messagebox.showerror("오류", str(e))

def create_filtered_file(original_file_path, duplicated_values):
    """중복된 수용가번호만 포함된 새 파일 생성"""
    try:
        # 원본 파일 읽기
        df = pd.read_excel(original_file_path, dtype=str).fillna('')
        
        # 중복된 수용가번호만 필터링
        filtered_df = df[df['수용가번호'].isin(duplicated_values)].copy()
        
        if filtered_df.empty:
            messagebox.showinfo("결과", "중복된 수용가번호가 없습니다.")
            return
        
        # 새 파일명 생성
        file_dir = os.path.dirname(original_file_path)
        file_name = os.path.basename(original_file_path)
        name, ext = os.path.splitext(file_name)
        new_file_path = os.path.join(file_dir, f"{name}_중복수용가만{ext}")
        
        # 파일 저장
        with pd.ExcelWriter(new_file_path, engine='openpyxl') as writer:
            filtered_df.to_excel(writer, index=False, sheet_name='중복수용가')
            
            # 중복 수용가번호 적색 표시
            wb = writer.book
            ws = wb['중복수용가']
            
            # 헤더에서 수용가번호 열 찾기
            headers = [cell.value for cell in ws[1]]
            col_idx = headers.index('수용가번호') + 1
            red_font = Font(color="FF0000")
            
            # 중복된 행들 적색 표시
            for row in range(2, ws.max_row + 1):
                cell = ws.cell(row=row, column=col_idx)
                cell.font = red_font
        
        messagebox.showinfo("완료", f"중복 수용가번호만 포함된 파일이 생성되었습니다:\n{new_file_path}")
        
    except Exception as e:
        messagebox.showerror("필터링 오류", str(e))

def read_google_sheet(sheet_url, credentials_path, worksheet_name):
    gc = gspread.service_account(filename=credentials_path)
    sh = gc.open_by_url(sheet_url)
    worksheet = sh.worksheet(worksheet_name)
    all_values = worksheet.get_all_values()
    header = all_values[0]
    idx_account = header.index('계정명')
    idx_service = header.index('서비스코드')
    idx_len = header.index('수용가번호길이')
    idx_struct = header.index('고객번호구조')
    data = [
        [row[idx_account], row[idx_service], row[idx_len], row[idx_struct]]
        for row in all_values[1:] if row[idx_account] and row[idx_service]
    ]
    df = pd.DataFrame(data, columns=['계정명', '서비스코드', '수용가번호길이', '고객번호구조'])
    return df

# GUI 구성
sheet_url = "https://docs.google.com/spreadsheets/d/10XO7o99fYr4e_I_etJF3_eCFa2_dsDs2egSWeu1GAls/edit#gid=679649875"
credentials_path = r"C:\제품등록\gcp9304-4410543fedf2.json"
worksheet_name = "IN형식"
df = read_google_sheet(sheet_url, credentials_path, worksheet_name)

window = tk.Tk()
window.title("임포트체커 + 통계 + 중복표시 (v250808)")
window.geometry("1000x600")

# 녹색 램프 프레임 (왼쪽 상단)
lamp_frame = tk.Frame(window)
lamp_frame.pack(anchor=tk.NW, padx=10, pady=(10, 0))

# 녹색 램프 (작은 네모박스)
lamp_canvas = tk.Canvas(lamp_frame, width=20, height=10, bg='white', highlightthickness=1, highlightbackground='black')
lamp_canvas.pack(side=tk.LEFT)

# 램프 상태 변수
lamp_on = True

def toggle_lamp():
    global lamp_on
    if lamp_on:
        lamp_canvas.configure(bg='lightgreen')
    else:
        lamp_canvas.configure(bg='white')
    lamp_on = not lamp_on
    # 1초마다 반복
    window.after(1000, toggle_lamp)

# 램프 시작
toggle_lamp()

site_frame = tk.Frame(window)
site_frame.pack(fill=tk.X, padx=10, pady=(10, 0))

site_label = tk.Label(site_frame, text="계정명 선택:")
site_label.pack(side=tk.LEFT)

exclude_accounts = ['나라장터', '농촌공사', '로우리스', '']
filtered_df = df[~df['계정명'].isin(exclude_accounts)].copy()
filtered_df = filtered_df.sort_values('계정명').reset_index(drop=True)
site_combobox = ttk.Combobox(site_frame, values=list(filtered_df['계정명']), state="readonly")
site_combobox.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(5, 0))

def on_site_select(event):
    idx = site_combobox.current()
    if idx >= 0:
        account_name = filtered_df.iloc[idx]['계정명']
        service_code = filtered_df.iloc[idx]['서비스코드']
        num_len = filtered_df.iloc[idx]['수용가번호길이']
        struct = filtered_df.iloc[idx]['고객번호구조']
        messagebox.showinfo("선택한 계정", f"계정명: {account_name}\n서비스코드: {service_code}\n수용가번호길이: {num_len}\n고객번호구조: {struct}")
site_combobox.bind('<<ComboboxSelected>>', on_site_select)

btn_sql = tk.Button(window, text="엑셀 파일 선택 및 SQL 변환", command=generate_sql_from_excel, bg="lightblue")
btn_sql.pack(pady=(10, 5))

btn_analyze = tk.Button(window, text="엑셀 파일 선택 및 수용가 통계 분석", command=analyze_excel_customer_stats, bg="lightgreen")
btn_analyze.pack(pady=(0, 5))

btn_dup = tk.Button(window, text="중복 수용가번호 적색 표시", command=mark_duplicates_in_place, bg="salmon")
btn_dup.pack(pady=(0, 10))

frame = tk.Frame(window)
frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

x_scrollbar = tk.Scrollbar(frame, orient=tk.HORIZONTAL)
x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

y_scrollbar = tk.Scrollbar(frame)
y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

result_text = tk.Text(frame, wrap=tk.NONE, xscrollcommand=x_scrollbar.set, yscrollcommand=y_scrollbar.set)
result_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

x_scrollbar.config(command=result_text.xview)
y_scrollbar.config(command=result_text.yview)

window.mainloop()
