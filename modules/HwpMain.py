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

    def sorting(self):
        if self.flag == 0:
            #--- 0 : num --- 1 : date --- 2 : product --- 3 : company --- 4 : classification --- 5 : name ---
            for idx, val in enumerate(os.listdir(self.directory)):
                try:
                    self.file_list.append((re.sub("컨설팅 지원 보고서_|\)|.hwp", "", val).replace("(","_")).split("_"))
                    self.file_list[idx].insert(0, f"A{idx+1}")
                    if self.file_list[idx][4] == "기술 검토":
                        self.cfc01+=1
                    elif self.file_list[idx][4] == "설계 및 제작":
                        self.cfc02+=1
                    self.file_list[idx][1] = datetime.strptime(('20' + self.file_list[idx][1]), '%Y%m%d').strftime('%Y-%m-%d')
                    self.file_list[idx][1], self.file_list[idx][3], self.file_list[idx][4] = self.file_list[idx][3], self.file_list[idx][4], self.file_list[idx][1]
                except Exception as e:
                    print(e)
                    print(val + " - 불필요한 파일입니다.")
                    pass
            self.summary = f"기술 검토 {self.cfc01}건 / 설계 및 제작 {self.cfc02}건"
            self.file_list.insert(0, ["순번", "기업명", "의뢰제품", "지원유형", "완료일자", "담당자"])

            return self.file_list
        elif self.flag == 1:
            print("flag is 1")

    def hwpInit(self):
        self.hwp = HwpFunctions(self.directory)
    def hwpPageSetup(self):
        self.hwp.PageSetup() #hwp 여백 설정
    def hwpHeadLine(self):
        self.hwp.HeadLine(len(self.file_list)-1, self.cfc01 + self.cfc02, self.summary) #헤드라인 문구 작성
    def hwpCreateChart(self):
        self.hwp.CreateChart(len(self.file_list) + 1)
    def hwpChartFill(self):
        self.hwp.ChartFill(self.file_list)
    def hwpInsertFile(self):
        self.hwp.InsertFile(self.directory)
        # hwp.XHwpWindows.Item(0).Visible = True

if __name__ == "__main__":
    app = HwpMain("0", "C:\\Users\\SHIN\\Documents\\모음")