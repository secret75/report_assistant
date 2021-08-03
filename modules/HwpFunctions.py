import win32com.client as win32
import os

class HwpFunctions:
    def __init__(self, directory):
        self.directory = directory 

        self.hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        self.hwp.RegisterModule("FilePathCheckDLL", "SecurityModule")
        self.hwp.XHwpWindows.Item(0).Visible = True

    def PageSetup(self):

        act = self.hwp.CreateAction("PageSetup")
        pset = act.CreateSet()
        act.GetDefault(pset)
        pset.SetItem("ApplyTo", 3)
        
        item_set = pset.CreateItemSet("PageDef", "PageDef")
        
        item_set.SetItem("TopMargin", self.hwp.MiliToHwpUnit(5.0))
        item_set.SetItem("BottomMargin", self.hwp.MiliToHwpUnit(5.0))
        item_set.SetItem("LeftMargin", self.hwp.MiliToHwpUnit(15.0))
        item_set.SetItem("RightMargin", self.hwp.MiliToHwpUnit(15.0))
        item_set.SetItem("HeaderLen", self.hwp.MiliToHwpUnit(10.0))
        item_set.SetItem("FooterLen", self.hwp.MiliToHwpUnit(10.0))
        item_set.SetItem("GutterLen", 0)

        act.Execute(pset)
        
    def CreateChart(self, _Rows, _fileList, cnt):
        if _Rows > 26:
            cnt += 1
            row = 26
            self.hwp.HAction.GetDefault("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)  # 표 생성 시작
            self.hwp.HParameterSet.HTableCreation.Rows = row # 행 갯수
            self.hwp.HParameterSet.HTableCreation.Cols = 6  # 열 갯수
            self.hwp.HParameterSet.HTableCreation.WidthType = 2  # 너비 지정(0:단에맞춤, 1:문단에맞춤, 2:임의값)
            self.hwp.HParameterSet.HTableCreation.HeightType = 1  # 높이 지정(0:자동, 1:임의값)
            self.hwp.HParameterSet.HTableCreation.WidthValue = self.hwp.MiliToHwpUnit(148.0)  # 표 너비
            self.hwp.HParameterSet.HTableCreation.HeightValue = self.hwp.MiliToHwpUnit(200)  # 표 높이
            self.hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 5)  # 열 5개 생성
            self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(0, self.hwp.MiliToHwpUnit(8.0))  # 1열
            self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(1, self.hwp.MiliToHwpUnit(32.0))  # 2열
            self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(2, self.hwp.MiliToHwpUnit(60.0))  # 3열
            self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(3, self.hwp.MiliToHwpUnit(22.0))  # 4열
            self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(4, self.hwp.MiliToHwpUnit(20.0))  # 5열
            self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(5, self.hwp.MiliToHwpUnit(15.0))  # 6열
            self.hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", 5)  # 행 5개 생성
            for i in range(row):
                self.hwp.HParameterSet.HTableCreation.RowHeight.SetItem(i, self.hwp.MiliToHwpUnit(8.0))  # 1행
            self.hwp.HParameterSet.HTableCreation.TableProperties.TreatAsChar = 1  # 글자처럼 취급
            self.hwp.HParameterSet.HTableCreation.TableProperties.Width = self.hwp.MiliToHwpUnit(148)  # 표 너비
            self.hwp.HAction.Execute("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)  # 위 코드 실행
        else:
            cnt += 1
            row = _Rows
            self.hwp.HAction.GetDefault("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)  # 표 생성 시작
            self.hwp.HParameterSet.HTableCreation.Rows = row # 행 갯수
            self.hwp.HParameterSet.HTableCreation.Cols = 6  # 열 갯수
            self.hwp.HParameterSet.HTableCreation.WidthType = 2  # 너비 지정(0:단에맞춤, 1:문단에맞춤, 2:임의값)
            self.hwp.HParameterSet.HTableCreation.HeightType = 1  # 높이 지정(0:자동, 1:임의값)
            self.hwp.HParameterSet.HTableCreation.WidthValue = self.hwp.MiliToHwpUnit(148.0)  # 표 너비
            self.hwp.HParameterSet.HTableCreation.HeightValue = self.hwp.MiliToHwpUnit(200)  # 표 높이
            self.hwp.HParameterSet.HTableCreation.CreateItemArray("ColWidth", 5)  # 열 5개 생성
            self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(0, self.hwp.MiliToHwpUnit(8.0))  # 1열
            self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(1, self.hwp.MiliToHwpUnit(32.0))  # 2열
            self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(2, self.hwp.MiliToHwpUnit(60.0))  # 3열
            self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(3, self.hwp.MiliToHwpUnit(22.0))  # 4열
            self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(4, self.hwp.MiliToHwpUnit(20.0))  # 5열
            self.hwp.HParameterSet.HTableCreation.ColWidth.SetItem(5, self.hwp.MiliToHwpUnit(15.0))  # 6열
            self.hwp.HParameterSet.HTableCreation.CreateItemArray("RowHeight", 5)  # 행 5개 생성
            for i in range(row):
                self.hwp.HParameterSet.HTableCreation.RowHeight.SetItem(i, self.hwp.MiliToHwpUnit(8.0))  # 1행
            self.hwp.HParameterSet.HTableCreation.TableProperties.TreatAsChar = 1  # 글자처럼 취급
            self.hwp.HParameterSet.HTableCreation.TableProperties.Width = self.hwp.MiliToHwpUnit(148)  # 표 너비
            self.hwp.HAction.Execute("TableCreate", self.hwp.HParameterSet.HTableCreation.HSet)  # 위 코드 실행
 
        for idx, i in enumerate(_fileList):
            if idx < 26:
                for j in i:
                    self.SetFont()
                    if cnt < 2 and idx == 0:
                        self.hwp.HAction.GetDefault("CellFill", self.hwp.HParameterSet.HCellBorderFill.HSet)
                        self.hwp.HParameterSet.HCellBorderFill.FillAttr.type = self.hwp.BrushType("NullBrush|WinBrush")
                        self.hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceColor = self.hwp.RGBColor(160, 190, 224)
                        self.hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushHatchColor = self.hwp.RGBColor(0, 0, 0)
                        self.hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceStyle = self.hwp.HatchStyle("None")
                        self.hwp.HParameterSet.HCellBorderFill.FillAttr.WindowsBrush = 1
                        self.hwp.HAction.Execute("CellFill",self.hwp.HParameterSet.HCellBorderFill.HSet)
                        self.hwp.Run("CharShapeBold")

                    self.hwp.Run("ParagraphShapeAlignCenter")
                    act = self.hwp.CreateAction("InsertText")
                    pset = act.CreateSet()
                    pset.SetItem("Text", j)
                    act.Execute(pset)
                    self.hwp.Run("TableRightCell")
        self.hwp.MovePos(3)
        self.hwp.HAction.Run("BreakPage")
        if _Rows > 27:
            _fileList = _fileList[26:]
            self.CreateChart(_Rows-27, _fileList, cnt)


    def SetFont(self):
        _font = "맑은 고딕"
        self.hwp.HAction.GetDefault("CharShape", self.hwp.HParameterSet.HCharShape.HSet)
        self.hwp.HParameterSet.HCharShape.FaceNameUser = _font
        self.hwp.HParameterSet.HCharShape.FontTypeUser = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.RatioUser = 100
        self.hwp.HParameterSet.HCharShape.SizeUser = 100
        self.hwp.HParameterSet.HCharShape.FaceNameSymbol = _font
        self.hwp.HParameterSet.HCharShape.FontTypeSymbol = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.RatioSymbol = 100
        self.hwp.HParameterSet.HCharShape.SizeSymbol = 100
        self.hwp.HParameterSet.HCharShape.FaceNameOther = _font
        self.hwp.HParameterSet.HCharShape.FontTypeOther = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.RatioOther = 100
        self.hwp.HParameterSet.HCharShape.SizeOther = 100
        self.hwp.HParameterSet.HCharShape.FaceNameJapanese = _font
        self.hwp.HParameterSet.HCharShape.FontTypeJapanese = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.RatioJapanese = 100
        self.hwp.HParameterSet.HCharShape.SizeJapanese = 100
        self.hwp.HParameterSet.HCharShape.FaceNameHanja = _font
        self.hwp.HParameterSet.HCharShape.FontTypeHanja = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.RatioHanja = 100
        self.hwp.HParameterSet.HCharShape.SizeHanja = 100
        self.hwp.HParameterSet.HCharShape.FaceNameLatin = _font
        self.hwp.HParameterSet.HCharShape.FontTypeLatin = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.RatioLatin = 100
        self.hwp.HParameterSet.HCharShape.SizeLatin = 100
        self.hwp.HParameterSet.HCharShape.FaceNameHangul = _font
        self.hwp.HParameterSet.HCharShape.FontTypeHangul = self.hwp.FontType("TTF")
        self.hwp.HParameterSet.HCharShape.RatioHangul = 100
        self.hwp.HParameterSet.HCharShape.SizeHangul = 100
        self.hwp.HParameterSet.HCharShape.Height = self.hwp.PointToHwpUnit(10.0)
        self.hwp.HAction.Execute("CharShape", self.hwp.HParameterSet.HCharShape.HSet)


    def HeadLine(self, TotalCnt, _Sum, Summary):
        act = self.hwp.CreateAction("InsertText")
        pset = act.CreateSet()

        char_shape = self.hwp.CharShape
        char_shape.SetItem("Height", 1200)
        self.hwp.CharShape = char_shape

        pset.SetItem("Text", "1. 제품 제작 컨설징 지원 (A)\r\n")
        act.Execute(pset)

        char_shape.SetItem("Height", 1000)
        self.hwp.CharShape = char_shape

        pset.SetItem("Text", f"(1) 총 건수 : {TotalCnt}건\r\n")
        act.Execute(pset)
        pset.SetItem("Text", f"(2) 내부 컨설팅 결과 내역 : {_Sum}건 ({Summary})\r\n")
        act.Execute(pset)

    def ChartFill(self, Chart_list):
        for idx, i in enumerate(Chart_list):
            for j in i:
                self.SetFont()
                if idx == 0:
                    self.hwp.HAction.GetDefault("CellFill", self.hwp.HParameterSet.HCellBorderFill.HSet)
                    self.hwp.HParameterSet.HCellBorderFill.FillAttr.type = self.hwp.BrushType("NullBrush|WinBrush")
                    self.hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceColor = self.hwp.RGBColor(160, 190, 224)
                    self.hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushHatchColor = self.hwp.RGBColor(0, 0, 0)
                    self.hwp.HParameterSet.HCellBorderFill.FillAttr.WinBrushFaceStyle = self.hwp.HatchStyle("None")
                    self.hwp.HParameterSet.HCellBorderFill.FillAttr.WindowsBrush = 1
                    self.hwp.HAction.Execute("CellFill",self.hwp.HParameterSet.HCellBorderFill.HSet)
                    self.hwp.Run("CharShapeBold")
                self.hwp.Run("ParagraphShapeAlignCenter")
                act = self.hwp.CreateAction("InsertText")
                pset = act.CreateSet()
                pset.SetItem("Text", j)
                act.Execute(pset)
                self.hwp.Run("TableRightCell")
        self.hwp.MovePos(3)
        self.hwp.HAction.Run("BreakPage")

    def InsertFile(self, path, flag):
        tmp = ""
        if flag == 0:
            tmp = "A"
        elif flag == 1:
            tmp = "B"
            
        file_list = os.listdir(path)
        field_list = []
        for idx, i in enumerate(file_list):
            self.hwp.HAction.GetDefault("InsertFile", self.hwp.HParameterSet.HInsertFile.HSet)
            self.hwp.HParameterSet.HInsertFile.filename = path + "/" + i
            self.hwp.HParameterSet.HInsertFile.KeepSection = 0
            self.hwp.HParameterSet.HInsertFile.KeepCharshape = 0
            self.hwp.HParameterSet.HInsertFile.KeepParashape = 0
            self.hwp.HParameterSet.HInsertFile.KeepStyle = 0
            self.hwp.HAction.Execute("InsertFile", self.hwp.HParameterSet.HInsertFile.HSet)
            self.hwp.PutFieldText(f'_Num{{{{{idx}}}}}', f'{tmp}{idx+1}')
            self.hwp.MovePos(3)
            self.hwp.HAction.Run("BreakPage")

if __name__ == "__main__":
    app = HwpMainFunction("test address")
