import platform

import pandas as pd


def main():
    os = platform.system()

    # excel_list = ['/Users/user/Desktop/엑셀파일/raw/장부내역.xlsx', '/Users/user/Desktop/엑셀파일/raw/정산거래처내역.xlsx']

    # 파일 경로 설정
    first_file_path = input("첫번째 엑셀 파일 경로를 입력하세요: ")
    second_file_path = input("두번째 엑셀 파일 경로를 입력하세요: ")

    # 엑셀 파일 읽기
    first_df = pd.read_excel(first_file_path)
    second_df = pd.read_excel(second_file_path)

    # 날짜 형식 바꾸기
    if '일자' in first_df.columns:
        first_df['일자'] = pd.to_datetime(first_df['일자'], format='%Y%m%d', errors='coerce').dt.strftime('%Y-%m-%d')

    if '정산일' in second_df.columns:
        second_df['정산일'] = pd.to_datetime(second_df['정산일'], errors='coerce').dt.strftime('%Y-%m-%d')

    # 거래처명에 공백 제거
    first_df['거래처명'] = first_df['거래처명'].str.replace(" ", "")
    second_df['거래처명'] = second_df['거래처명'].str.replace(" ", "")

    # 거래처명을 기준으로 차이를 찾기 위해 DataFrame에서 거래처명만 추출
    first_set = set(first_df['거래처명'])
    second_set = set(second_df['거래처명'])

    # 두 데이터프레임 간의 거래처명 차이 찾기
    first_result = first_set - second_set
    second_result = second_set - first_set

    # 차이점이 있는 거래처명에 해당하는 데이터 추출
    first_result_df = first_df[first_df['거래처명'].isin(first_result)]
    second_result_df = second_df[second_df['거래처명'].isin(second_result)]

    # 결과 출력
    print("장부내역 only 결과:")
    print(first_result_df)

    print("정산거래처내역 only 결과:")
    print(second_result_df)

    if os == "Windows":
        with pd.ExcelWriter(r'C:\Users\USER\Desktop\결과\결과.xlsx') as writer:
            first_result_df.to_excel(writer, sheet_name='정산거래처내역에 없는 케이스', index=False)
            second_result_df.to_excel(writer, sheet_name='장부내역에 없는 케이스', index=False)
    else:
        with pd.ExcelWriter('/Users/user/Desktop/엑셀파일/result/결과.xlsx') as writer:
            first_result_df.to_excel(writer, sheet_name='정산거래처내역에 없는 케이스', index=False)
            second_result_df.to_excel(writer, sheet_name='장부내역에 없는 케이스', index=False)
    print("결과 파일이 생성되었습니다!")

if __name__ == '__main__':
    main()