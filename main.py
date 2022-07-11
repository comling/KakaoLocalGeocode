####################################################
# KakaoLocal Geocode용 프로그램
# Revision : 220711_0.1
# made by : Wonchul
####################################################

import os
import sys
import pandas as pd
from PyKakao import KakaoLocal

if getattr(sys, 'frozen', False):
    program_directory = os.path.dirname(os.path.abspath(sys.executable))
else:
    program_directory = os.path.dirname(os.path.abspath(__file__))

KAKAO_REST_API_KEY = 'cd24225fcfcd41a9094ba4b84277d01b'

def kakao_address(address):
    ### 카카오로컬API 라이브러리를 사용한 주소 검색 조회
    KL = KakaoLocal(KAKAO_REST_API_KEY)

    ### 카카오 API 로컬주소 확인
    address_get = KL.search_address(address)
    ### 주소 정보 저장
    address_info = []
    result = []
    temp = ''

    try:
        address_info.append(address)
        address_info.append(address_get['documents'][0]['address']['address_name'])
        try:
            address_info.append(address_get['documents'][0]['road_address']['address_name'])
        except TypeError:
            address_info.append(address_get['documents'][0]['address']['address_name'][0:2])

        for i in address_info:
            if ('서울' in i) == True:
                temp = i.replace('서울', '서울특별시')
            elif ('부산' in i) == True:
                temp = i.replace('부산', '부산광역시')
            elif ('대구' in i) == True:
                temp = i.replace('대구', '대구광역시')
            elif ('인천' in i) == True:
                temp = i.replace('인천', '인천광역시')
            elif ('광주' in i) == True:
                temp = i.replace('광주', '광주광역시')
            elif ('대전' in i) == True:
                temp = i.replace('대전', '대전광역시')
            elif ('울산' in i) == True:
                temp = i.replace('울산', '울산광역시')
            elif ('세종' in i) == True:
                temp = i.replace('세종', '세종특별자치시')
            elif ('경기' in i) == True:
                temp = i.replace('경기', '경기도')
            elif ('강원' in i) == True:
                temp = i.replace('강원', '강원도')
            elif ('충북' in i) == True:
                temp = i.replace('충북', '충청북도')
            elif ('충남' in i) == True:
                temp = i.replace('충남', '충청남도')
            elif ('전북' in i) == True:
                temp = i.replace('전북', '전라북도')
            elif ('전남' in i) == True:
                temp = i.replace('전남', '전라남도')
            elif ('경북' in i) == True:
                temp = i.replace('경북', '경상북도')
            elif ('경남' in i) == True:
                temp = i.replace('경남', '경상남도')
            elif ('제주' in i) == True:
                temp = i.replace('제주', '제주특별자치도')

            result.append(temp)

        result.append(address_get['documents'][0]['address']['x'])
        result.append(address_get['documents'][0]['address']['y'])

        return result

    except Exception as e:
        return [address, '', '', '', '']

if __name__ == '__main__':

    account_info_file = 'address.xlsx'
    account_info_data = pd.ExcelFile(os.path.join(program_directory, account_info_file))

    df_account_info = account_info_data.parse()
    account_data = df_account_info

    address_value = account_data['주소']
    address_result = []
    for i in range(0, len(address_value)):
        print('\r' + str(i) + '/' + str(len(address_value)), end="")
        address_result.append(kakao_address(address_value[i]))
    address_result_data = pd.DataFrame(address_result)
    df_address_result = pd.concat([account_data, address_result_data], axis=1)
    file_path = os.path.join(program_directory, 'address_geocode.xlsx')
    excel_writer = pd.ExcelWriter(file_path, engine='openpyxl')
    df_address_result.to_excel(excel_writer, index=None)
    excel_writer.save()