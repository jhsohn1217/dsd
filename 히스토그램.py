import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog

# 한글 폰트 설정 (koreanize-matplotlib 사용)
try:
    import koreanize_matplotlib
except ImportError:
    print("koreanize_matplotlib가 설치되지 않았습니다. 설치 후 실행해주세요.")

def select_file():
    """사용자 파일 선택 (GUI 방식)"""
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(
        title="엑셀 파일 선택",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    return file_path

def plot_histogram(file_path):
    """총매출 히스토그램 시각화"""
    try:
        xls = pd.ExcelFile(file_path)
        if '바차트_히스토그램' not in xls.sheet_names:
            raise ValueError("'바차트_히스토그램' 시트를 찾을 수 없습니다.")

        df = xls.parse('바차트_히스토그램')
        if '총 매출' not in df.columns:
            raise ValueError("'총 매출' 열이 존재하지 않습니다.")

        plt.figure(figsize=(10, 6))
        n, bins, patches = plt.hist(df['총 매출'], bins=8, edgecolor='black')

        for i in range(len(patches)):
            height = n[i]
            bin_center = 0.5 * (bins[i] + bins[i+1])
            plt.text(bin_center, height + 0.3, f'{int(height)}', ha='center', va='bottom')

        plt.title('총 매출 히스토그램 (8개 구간)')
        plt.xlabel('총 매출')
        plt.ylabel('빈도수')
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.tight_layout()
        plt.show()

    except Exception as e:
        print(f"[오류] {e}")

if __name__ == "__main__":
    print("🔍 히스토그램 생성기 시작...")
    file_path = select_file()
    if file_path:
        plot_histogram(file_path)
    else:
        print("파일이 선택되지 않았습니다.")
