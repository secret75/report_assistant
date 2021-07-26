import sys

from modules import *

global _dir

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setupUi()

        self.show()

    def setupUi(self):
        #title
        self.setWindowTitle("Report Assistant")
        self.setGeometry(700, 450, 470, 620)
        self.setMinimumSize(470, 620)
        self.setMaximumSize(470, 620)



        self.btn = QPushButton(self)
        self.btn.setGeometry(370, 10, 91, 31)
        self.btn.setText("폴더 열기")
        self.btn.clicked.connect(self.open_folder)

        self.execute_btn = QPushButton(self)
        self.execute_btn.setGeometry(10, 50, 451, 61)
        self.execute_btn.setText("시작")
        self.execute_btn.clicked.connect(self.hwp_sorting)

        self.line_dir = QLineEdit(self)
        self.line_dir.setGeometry(12, 11, 351, 31)
        self.line_dir.returnPressed.connect(self.save_directory)

        #TextBrowser set
        self.tb = QTextBrowser(self)
        self.tb.setGeometry(10, 120, 451, 491)

    
    def open_folder(self):
        global _dir
        _dir = QFileDialog.getExistingDirectory(None, '폴더 선택', 'C:/', QFileDialog.ShowDirsOnly)
        self.line_dir.setText(_dir)

    def save_directory(self):
        _dir = self.line_dir.text()

    def hwp_sorting(self):
        try:
            self.tb.append("==========문서 디렉토리 위치==========")
            self.tb.append(f"Open Directory : {_dir}")
            hwpframe = HwpMain("0", _dir)

            self.tb.append("==========문서 리스트 읽기==========")
            val = hwpframe.sorting()
            for i in val:
                self.tb.append('_'.join(i))

            self.tb.append("==========한글 객체 생성==========")
            hwpframe.hwpInit()

            self.tb.append("==========한글 문서 초기화==========")
            hwpframe.hwpPageSetup()

            self.tb.append("==========문서 헤드라인==========")
            hwpframe.hwpHeadLine()

            self.tb.append("==========표 그리기==========")
            hwpframe.hwpCreateChart()

            self.tb.append("==========보고서 목차 입력==========")
            hwpframe.hwpChartFill()

            self.tb.append("==========보고서 내용 복사==========")
            hwpframe.hwpInsertFile()

        except Exception as e:
            self.tb.append("문서 폴더를 지정해주세요.")
            self.tb.append(e)
        


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setWindowIcon(QIcon("CI_CLK_R1.png"))
    win = MainWindow()
    sys.exit(app.exec_())
