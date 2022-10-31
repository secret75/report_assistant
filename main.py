import sys

from modules import *

global _dir
global flag

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi()

        self.show()

    def setupUi(self):
        #title
        self.setWindowTitle("Report Assistant v0.2")
        self.setGeometry(400, 250, 470, 620)
        self.setMinimumSize(470, 620)
        self.setMaximumSize(470, 620)


        groupBox = QGroupBox('보고서 분류', self)
        groupBox.setGeometry(10, 5, 450, 40)

        self.radio1 = QRadioButton(self)
        self.radio1.setText("A. 컨설팅 지원 보고서")
        self.radio1.setGeometry(20, 20, 150, 20)
        self.radio1.clicked.connect(self.radioButtonClicked)
        self.radio2 = QRadioButton(self)
        self.radio2.setText("B. 시제품 제작 지원 보고서")
        self.radio2.setGeometry(250, 20, 200, 20)
        self.radio2.clicked.connect(self.radioButtonClicked)

        self.line_dir = QLineEdit(self)
        self.line_dir.setGeometry(10, 50, 350, 30)
        self.line_dir.returnPressed.connect(self.save_directory)

        self.btn = QPushButton(self)
        self.btn.setGeometry(365, 49, 96, 32)
        self.btn.setText("폴더 열기")
        self.btn.clicked.connect(self.open_folder)


        self.execute_btn = QPushButton(self)
        self.execute_btn.setGeometry(10, 90, 451, 61)
        self.execute_btn.setText("시작")
        self.execute_btn.clicked.connect(self.hwp_sorting)

        #TextBrowser set
        self.tb = QTextBrowser(self)
        self.tb.setGeometry(10, 160, 451, 451)
    def radioButtonClicked(self):
        global flag
        if self.radio1.isChecked():
            flag = 0
        elif self.radio2.isChecked():
            flag = 1
    
    def open_folder(self):
        global _dir
        _dir = QFileDialog.getExistingDirectory(None, '폴더 선택', 'C:/', QFileDialog.ShowDirsOnly)
        self.line_dir.setText(_dir)

    def save_directory(self):
        _dir = self.line_dir.text()

    def hwp_sorting(self):
        _dir = self.line_dir.text()
        try:

            self.tb.append("==========문서 디렉토리 위치==========")
            self.tb.append(f"Open Directory : {_dir}")
            hwpframe = HwpMain(flag, _dir)

            self.tb.append("==========한글 객체 생성==========")
            hwpframe.hwpInit(False)

            self.tb.append("==========문서 리스트 읽기==========")
            val = hwpframe.sorting()
            for i in val:
                self.tb.append('_'.join(i))


            self.tb.append("==========한글 객체 생성==========")
            hwpframe.hwpClear()
            hwpframe.hwpInit(True)

            self.tb.append("==========한글 문서 초기화==========")
            hwpframe.hwpPageSetup()

            self.tb.append("==========문서 헤드라인==========")
            hwpframe.hwpHeadLine()

            self.tb.append("==========표 그리기==========")
            hwpframe.hwpCreateChart()

            # self.tb.append("==========보고서 목차 입력==========")
            # hwpframe.hwpChartFill()

            self.tb.append("==========보고서 내용 복사==========")
            hwpframe.hwpInsertFile()

        except Exception as e:
            self.tb.append("에러 발생!!")
            self.tb.append(e)
        


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("clk.ico"))
    win = MainWindow()
    sys.exit(app.exec_())
