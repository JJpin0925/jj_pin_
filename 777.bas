Attribute VB_Name = "Module1"
Sub asc()
Attribute asc.VB_ProcData.VB_Invoke_Func = "z\n14"
'
' asc ����
'
' �ֳt��: Ctrl+z
'
    Range("B1").Select
    Application.WindowState = xlMinimized
    Application.WindowState = xlNormal
    Application.WindowState = xlMinimized
    Application.WindowState = xlNormal
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Application.WindowState = xlMinimized
    Application.WindowState = xlNormal
End Sub
Sub acc()
Attribute acc.VB_ProcData.VB_Invoke_Func = "x\n14"
'
' acc ����
'
' �ֳt��: Ctrl+x
'
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("�u�@��1").Sort.SortFields.Add Key:=Range("B2:B414"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("�u�@��1").Sort
        .SetRange Range("A1:B414")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
