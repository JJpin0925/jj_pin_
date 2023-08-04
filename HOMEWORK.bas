Sub CopyAndSum()
'
' CopyAndSum 巨集
'

'新增工作表
    Sheets.Add After:=ActiveSheet
    Sheets("工作表1").Select
'改名為總人數統計
    Sheets("工作表1").Name = "總人數統計"
'點選國民出國目的地人數統計2020工作表
    Sheets("國民出國目的地人數統計2020").Select
'複製範圍內文字
    Range("A1:B36").Select
    Selection.Copy
'貼上到總人數統計
    Sheets("總人數統計").Select
    ActiveSheet.Paste
    Columns("A:A").EntireColumn.AutoFit
    Columns("B:B").EntireColumn.AutoFit
'點選C1儲存格改名為總和
    Range("C1").Select
    ActiveCell.FormulaR1C1 = "總和"

'從excel第2列到36列
rowNum = Cells(Rows.Count, 1).End(xlUp).Row
For rowIdx = 2 To rowNum
'第4張表的c欄位(總數)=第一張表的c欄位(數量)+第二張表的c欄位(數量)+第二張表的c欄位(數量)
Sheets(4).Cells(rowIdx, 3).Value = Sheets(1).Cells(rowIdx, 3).Value + Sheets(2).Cells(rowIdx, 3).Value + Sheets(3).Cells(rowIdx, 3).Value
 

Next

' 排序功能產出 巨集
'

'
    ActiveSheet.Buttons.Add(384.9, 28.2, 114.9, 30.3).Select
    Application.CutCopyMode = False
    Selection.OnAction = "出國人口大到小排序"
    Selection.Characters.Text = "大至小排序"
    With Selection.Characters(Start:=1, Length:=5).Font
        .Name = "新細明體"
        .FontStyle = "標準"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("I4").Select
    ActiveSheet.Buttons.Add(386.1, 73.2, 112.5, 26.1).Select
    Selection.OnAction = "出國人口小到大排序"
    Selection.Characters.Text = "小至大排序"
    With Selection.Characters(Start:=1, Length:=5).Font
        .Name = "新細明體"
        .FontStyle = "標準"
        .Size = 12
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ColorIndex = 1
    End With
    Range("J6").Select

End Sub

Sub 出國人口小到大排序()
'
' 人數小到大 巨集
'
' 只用在總人數統計工作表
'
    Range("C1").Select
    ActiveWorkbook.Worksheets("總人數統計").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("總人數統計").Sort.SortFields.Add2 Key:=Range("C2:C36") _
        , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("總人數統計").Sort
        .SetRange Range("A1:C36")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
End With

End Sub

Sub 出國人口大到小排序()
'
' 人數大到小 巨集
'
' 只用在總人數統計工作表
'
    Range("C1").Select
    ActiveWorkbook.Worksheets("總人數統計").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("總人數統計").Sort.SortFields.Add2 Key:=Range("C2:C36") _
        , SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("總人數統計").Sort
        .SetRange Range("A1:C36")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
End With
  
    
End Sub




Private Sub btnCustRange_Click()
Dim rangAddress As String
rangAddress = InputBox("請輸入儲存格位址")
Dim ChartType As Integer
Select Case combCType.Value
  Case "圓餅圖"
   ChartType = 5
  Case "橫條圖"
   ChartType = 57
  Case "直條圖"
   ChartType = 51
  Case "折線圖"
   ChartType = 4
End Select
Sheets(combSName.Value).Activate '啟動第i張工作表
Dim dtRange As Range '宣告範圍1變數
Set dtRange = ActiveSheet.UsedRange '範圍1設定為已使用區域
ActiveSheet.Shapes.AddChart2(201, ChartType).Select '新增圖表
ActiveChart.SetSourceData Source:=Range(rangAddress) '圖表來源
End Sub

Private Sub btnPrint_Click()
MsgBox (combSName.Value)
Sheets(combSName.Value).PrintOut
End Sub

Private Sub btnRun_Click()
Dim ChartType As Integer
Select Case combCType.Value
  Case "圓餅圖"
   ChartType = 5
  Case "橫條圖"
   ChartType = 57
  Case "直條圖"
   ChartType = 51

End Select

For i = 1 To Sheets.Count
Sheets(i).Activate '啟動第i張工作表
Dim dtRange As Range '宣告範圍1變數
Set dtRange = ActiveSheet.UsedRange '範圍1設定為已使用區域
ActiveSheet.Shapes.AddChart2(201, ChartType).Select '新增圖表
ActiveChart.SetSourceData Source:=dtRange '圖表來源
Next
MsgBox ("所有工作表已繪製完成")
End Sub

Private Sub btnSheet_Click()
Dim ChartType As Integer
Select Case combCType.Value
  Case "圓餅圖"
   ChartType = 5
  Case "橫條圖"
   ChartType = 57
  Case "直條圖"
   ChartType = 51
End Select
Sheets(combSName.Value).Activate '啟動第i張工作表
Dim dtRange As Range '宣告範圍1變數
Set dtRange = ActiveSheet.UsedRange '範圍1設定為已使用區域
ActiveSheet.Shapes.AddChart2(201, ChartType).Select '新增圖表
ActiveChart.SetSourceData Source:=dtRange '圖表來源

End Sub

Private Sub UserForm_Initialize()
combCType.AddItem "圓餅圖"
combCType.AddItem "直條圖"
combCType.AddItem "橫條圖"

Dim i As Integer
For i = 1 To Sheets.Count
combSName.AddItem Sheets(i).Name
Next

End Sub
