import os
import pyautogui
import sys
import pandas as pd
import webbrowser
import win32com.client

####엑셀 파일 읽어오기
# 배송리스트가 담긴 폴더 읽어오기
delivery_list_folder_name = "delivery_list"
if os.path.isdir(delivery_list_folder_name):#폴더 있는지 확인
    delivery_list_folder = os.listdir(delivery_list_folder_name)
else:
    os.makedirs(delivery_list_folder_name)

# 배송리스트 파일 읽어오기
if len(delivery_list_folder) != 0:
    delivery_list = pd.read_csv(f"{delivery_list_folder_name}\\{delivery_list_folder[0]}", encoding='utf-8')
else:
    pyautogui.alert(f"{delivery_list_folder_name} 폴더에 파일이 없습니다!", button="프로그램 종료")
    sys.exit() #프로그램 종료


####카페24 양식에 맞게 수정
try:
    # C열의 데이터까지만 남겨두기.
    upload_to_cafe24 = delivery_list.iloc[:, :3]

    # D1 셀에 "수량" 추가 
    upload_to_cafe24.insert(3, '수량', '') 

    # 수정된 내용을 새로운 CSV 파일로 저장
    upload_to_cafe24.to_csv("excel_sample_old.csv", index=False, encoding='utf-8-sig')
 
except Exception as e:
    print(f"파일 편집 중 오류가 발생했습니다: {e}")

####매크로 실행
#프린터 지정
#매크로 실행
try:
    # 엑셀 애플리케이션 시작
    excel = win32com.client.Dispatch("Excel.Application")
    
    # 엑셀 애플리케이션을 표시하지 않음 (백그라운드에서 실행)
    excel.Visible = False
    
    
    # 엑셀 애플리케이션 종료
    excel.Quit()
    
  #  print(f"'{macro_name}' 매크로가 성공적으로 실행되었습니다.")
    
except Exception as e:
    print(f"실행 중 오류가 발생했습니다: {e}")


####한진택배 사이트 열기
webbrowser.open("https://focus.hanjin.com/login")
