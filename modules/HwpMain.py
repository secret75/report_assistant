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

    def sorting(self):
        if self.flag == 0:
            #--- 0 : num --- 1 : date --- 2 : product --- 3 : company --- 4 : classification --- 5 : name ---
            dic = {}
            for idx, val in enumerate(os.listdir(self.directory)):
                try:
                    self.file_list.append((re.sub("컨설팅 지원 보고서_|\)|.hwp", "", val).replace("(","_")).split("_"))
                    self.file_list[idx].insert(0, f"A{idx+1}")
                    if self.file_list[idx][4] not in dic:
                        dic[self.file_list[idx][4]] = 1
                    else:
                        dic[self.file_list[idx][4]] += 1
                    self.file_list[idx][1] = datetime.strptime(('20' + self.file_list[idx][1]), '%Y%m%d').strftime('%Y-%m-%d')
                    self.file_list[idx][1], self.file_list[idx][3], self.file_list[idx][4] = self.file_list[idx][3], self.file_list[idx][4], self.file_list[idx][1]
                except Exception as e:
                    print(e)
                    print(val + " - 불필요한 파일입니다.")
                    pass
            a = []
            for i in dic:
                a.append(f"{i} {dic[i]}건")
                self._sum += dic[i]
            self.summary = ' / '.join(a)
            self.file_list.insert(0, ["순번", "기업명", "의뢰제품", "지원유형", "완료일자", "담당자"])

            return self.file_list
        elif self.flag == 1:
            dic = {}
            for idx, val in enumerate(os.listdir(self.directory)):
                try:
                    self.file_list.append((re.sub("제품화 지원 보고서_|\)|.hwp", "", val).replace("(","_")).split("_"))
                    self.file_list[idx].insert(0, f"B{idx+1}")

                    a = self.file_list[idx][4].replace(" ", "")
                    if a not in dic:
                        dic[a] = 1
                    else:
                        dic[a] += 1
                    self.file_list[idx][1] = datetime.strptime(('20' + self.file_list[idx][1]), '%Y%m%d').strftime('%Y-%m-%d')
                    self.file_list[idx][1], self.file_list[idx][3], self.file_list[idx][4] = self.file_list[idx][3], self.file_list[idx][4], self.file_list[idx][1]
                except Exception as e:
                    print(e)
                    print(val + " - 불필요한 파일입니다.")
            a = []
            for i in dic:
                a.append(f"{i} : {dic[i]}건")
                self._sum += dic[i]
            self.summary = ' / '.join(a)
            self.file_list.insert(0, ["순번", "기업명", "의뢰제품", "지원유형", "완료일자", "담당자"])

            return self.file_list

    def hwpInit(self):
        self.hwp = HwpFunctions(self.directory)
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

if __name__ == "__main__":
    app = HwpMain("0", "C:\\Users\\SHIN\\Documents\\모음")