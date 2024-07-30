import platform
import pandas as pd
from tkinter import Tk
from tkinter.filedialog import askopenfilename

def main():
    os = platform.system()

    # 파일 선택 창 띄우기
    def select_file(prompt):
        Tk().withdraw()  # 기본 tkinter 창을 숨김
        filename = askopenfilename(title=prompt, filetypes=[("Excel files", "*.xlsx *.xls")])
        return filename

    # 첫 번째 엑셀 파일 선택
    first_file_path = select_file("첫 번째 엑셀 파일을 선택하세요:")
    if not first_file_path:
        print("첫 번째 파일이 선택되지 않았습니다.")
        exit()

    # 두 번째 엑셀 파일 선택
    second_file_path = select_file("두 번째 엑셀 파일을 선택하세요:")
    if not second_file_path:
        print("두 번째 파일이 선택되지 않았습니다.")
        exit()

    # 파일 경로 설정
    # first_file_path = input("첫번째 엑셀 파일 경로를 입력하세요: ")
    # second_file_path = input("두번째 엑셀 파일 경로를 입력하세요: ")
    # first_file_path = '/Users/user/Desktop/장부내역.xlsx'
    # second_file_path = '/Users/user/Desktop/정산거래처내역.xlsx'

    # 엑셀 파일 읽기
    first_df = pd.read_excel(first_file_path)
    second_df = pd.read_excel(second_file_path)

    # 날짜 형식 바꾸기
    # if '일자' in first_df.columns:
    #     first_df['일자'] = pd.to_datetime(first_df['일자'], format='%Y%m%d', errors='coerce').dt.strftime('%Y-%m-%d')

    if '정산일' in second_df.columns:
        second_df['정산일'] = pd.to_datetime(second_df['정산일'], errors='coerce').dt.strftime('%Y-%m-%d')

    # 거래처명에 공백 제거
    first_df['거래처명'] = first_df['거래처명'].str.replace(" ", "")
    second_df['거래처명'] = second_df['거래처명'].str.replace(" ", "")

    # 각 엑셀 데이터 출력
    print("원본 데이터:")
    print(first_df)

    print("원본 데이터:")
    print(second_df)

    # 거래처명을 기준으로 차이를 찾기 위해 DataFrame에서 거래처명만 추출
    first_set = set(first_df['거래처명'])
    second_set = set(second_df['거래처명'])

    # 두 데이터프레임 간의 거래처명 차이 찾기
    first_result = first_set - second_set
    second_result = second_set - first_set

    # 차이점이 있는 거래처명에 해당하는 데이터 추출
    first_result_df = first_df[first_df['거래처명'].isin(first_result)]
    second_result_df = second_df[second_df['거래처명'].isin(second_result)]

    # 날짜를 기준으로 정렬
    if '일자' in first_result_df.columns:
        first_result_df = first_result_df.sort_values(by='일자')
    if '정산일' in second_result_df.columns:
        second_result_df = second_result_df.sort_values(by='정산일')

    # 결과 파일 저장
    if os == "Windows":
        output_path = r'C:\Users\USER\Desktop\result\결과.xlsx'
    else:
        output_path = '/Users/user/Desktop/excel/result/결과.xlsx'

    with pd.ExcelWriter(output_path) as writer:
        first_result_df.to_excel(writer, sheet_name='정산거래처내역에 없는 케이스', index=False)
        second_result_df.to_excel(writer, sheet_name='장부내역에 없는 케이스', index=False)

    print("결과 파일이 생성되었습니다!")

if __name__ == '__main__':
    main()