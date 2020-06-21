# Aouthor Beomsang Kim
# Produced 21-Jun-2020
# Origin netScout
# Purpose 엑셀 리포트 자동화 실행
#         지난주와 이번주 리포트의 날짜를 입력받음
# Comment MVC페턴으로 만들려 했는데 MVC가 아님 잉잉

import control

print("양식 비교를 위함입니다")
print("지난주 리포트 생성 날자를 입력해 주세요")
lastWeekDate = input("YYYYMMDD: ")
print("이번주 리포트 생성 날자를 입력해 주세요")
thisWeekDate = input("YYYYMMDD: ")
print("잠시만 기다려 주세요")
control.controling()
print(thisWeekDate + " 일자 리포트 생성이 완료되었습니다")