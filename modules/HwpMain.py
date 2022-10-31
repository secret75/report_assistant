import re
import os
from datetime import datetime
if __name__ == "__main__":
    from HwpFunctions import *
else:
    from . HwpFunctions import *



class HwpMain():
    def __init__(self, flag, directory):
        self.flag = flag
        self.directory = directory

        self.file_list = []
        self.cfc01 = 0
        self.cfc02 = 0
        self.summary = ""
        self._sum = 0

    def classification(self, title, prefix):
        #--- 0 : num --- 1 : date --- 2 : product --- 3 : company --- 4 : classification --- 5 : name ---
        dic = {}
        file_list = []
        for idx, val in enumerate(os.listdir(self.directory)):
            file_list.append((re.sub(f"{title}_|\)|.hwp", "", val).replace("(","_")).split("_")) # 파일 이름 간소화 및 "_"로 나누기
            file_list[idx].insert(0, f"{prefix}{idx+1}") # 각 항목 0번째 A1~n prefix 부여
            a = file_list[idx][4].replace(" ", "") # 분야 항목에 공백 삭제
            if a not in dic: # 딕셔너리에 추가
                dic[a] = 1
            else:
                dic[a] += 1
            file_list[idx][1] = datetime.strptime(('20' + file_list[idx][1]), '%Y%m%d').strftime('%Y-%m-%d') # 날짜 형식 변경
            file_list[idx][1], file_list[idx][3], file_list[idx][4] = file_list[idx][3], file_list[idx][4], file_list[idx][1]
        file_list.insert(0, ["순번", "기업명", "의뢰제품", "지원유형", "완료일자", "담당자"])
        print(dic)
        print(type(dic))

        a = []
        for i in dic:
            a.append(f"{i} {dic[i]}건")
            self._sum += dic[i]
        self.summary = " / ".join(a)

            
        return file_list, dic

    def sorting(self):
        fieldValue = self.hwp.getFieldValue()
        for i in fieldValue:
            print(i)

        if self.flag == 0:
            title = "컨설팅 지원 보고서"
            prefix = "A"
            file_list, dic = self.classification(title, prefix)
            self.file_list = file_list

            return self.file_list

        elif self.flag == 1:
            title = "시제품 제작 지원 보고서"
            prefix = "B"
            file_list, dic = self.classification(title, prefix)
            self.file_list = file_list

            return self.file_list

    def hwpInit(self, visible):
        self.hwp = HwpFunctions(self.directory, visible)
    def hwpPageSetup(self):
        self.hwp.PageSetup() #hwp 여백 설정
    def hwpHeadLine(self):
        self.hwp.HeadLine(len(self.file_list)-1, self._sum, self.summary) #헤드라인 문구 작성
    def hwpCreateChart(self):
        cnt = 0
        self.hwp.CreateChart(len(self.file_list) + 1, self.file_list, cnt)
    # def hwpChartFill(self):
    #     self.hwp.ChartFill(self.file_list)
    def hwpInsertFile(self):
        self.hwp.InsertFile(self.directory, self.flag)
        # hwp.XHwpWindows.Item(0).Visible = True
    def hwpClear(self):
        self.hwp.HwpClear()

if __name__ == "__main__":
    app = HwpMain("0", "C:/Users/SHIN/Desktop/모음")
