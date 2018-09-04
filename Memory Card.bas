Sub CreateDB()
    
    '这个表格包括Origin, Jan, Feb,March,April,May,June,July,Aug,Sep,Oct,Nov,Dec共计13个sheets
    'Orign第1行为单词行，第2行至第9行是时间（1,3,7,14,21,28,60,90 days)
    '将文件命名为Memory Card.xlsm

    Windows("Memory Card.xlsm").Activate
    
    '
    ' Define the 12 sheets to the 12 months
    '
    Dim sht_jan As Worksheet
    Dim sht_feb As Worksheet
    Dim sht_march As Worksheet
    Dim sht_april As Worksheet
    Dim sht_may As Worksheet
    Dim sht_june As Worksheet
    Dim sht_july As Worksheet
    Dim sht_aug As Worksheet
    Dim sht_sep As Worksheet
    Dim sht_oct As Worksheet
    Dim sht_nov As Worksheet
    Dim sht_dec As Worksheet
    Set sht_jan = Application.ThisWorkbook.Worksheets("Jan")
    Set sht_feb = Application.ThisWorkbook.Worksheets("Feb")
    Set sht_march = Application.ThisWorkbook.Worksheets("March")
    Set sht_april = Application.ThisWorkbook.Worksheets("April")
    Set sht_may = Application.ThisWorkbook.Worksheets("May")
    Set sht_june = Application.ThisWorkbook.Worksheets("June")
    Set sht_july = Application.ThisWorkbook.Worksheets("July")
    Set sht_aug = Application.ThisWorkbook.Worksheets("Aug")
    Set sht_sep = Application.ThisWorkbook.Worksheets("Sep")
    Set sht_oct = Application.ThisWorkbook.Worksheets("Oct")
    Set sht_nov = Application.ThisWorkbook.Worksheets("Nov")
    Set sht_dec = Application.ThisWorkbook.Worksheets("Dec")
    
    '
    ' Define the sheet named orgin
    
    Set sht_origin = Application.ThisWorkbook.Worksheets("Origin")
    
    '
    ' 构建Origin
    '
    Dim ACI As Integer
    Dim JSRow, JSColumn As Integer
    Dim JiSuanRiQiDB As Range
    
    For ACI = 1 To 10
        Cells((Cells(65535, 3).End(xlUp).Row + 1), 2).Select
        Set JiSuanRiQiDB = Selection
        If JiSuanRiQiDB = "" Then
                Exit For
        End If
    
        JSRow = JiSuanRiQiDB.Row
        JSColumn = JiSuanRiQiDB.Column
        
        Cells(JSRow, JSColumn + 1).Value = Selection + 1
        Cells(JSRow, JSColumn + 2).Value = Selection + 3
        Cells(JSRow, JSColumn + 3).Value = Selection + 7
        Cells(JSRow, JSColumn + 4).Value = Selection + 14
        Cells(JSRow, JSColumn + 5).Value = Selection + 21
        Cells(JSRow, JSColumn + 6).Value = Selection + 28
        Cells(JSRow, JSColumn + 7).Value = Selection + 60
        Cells(JSRow, JSColumn + 8).Value = Selection + 90
    Next ACI
    
    '
    ' i 用于标记row ， j 用于标记column
    Dim i, j As Integer
    
    
    '
    ' copy_time_rng 用于存储时间值
    ' copy_word_rng 用于存储copy_time_rng对应的单词
    '
    Dim copy_time_rng As Range
    Dim copy_word_rng As Range
    
    '
    '  定义参数tian，yue，用于存储day， month的数据
    '
    Dim tian, yue As Integer

    
    ' 对时间区域的数据进行全匹配
    
    ' i 用于标记row
    For i = 2 To 5000
    
        ' j 用于标记column
        For j = 1 To 9
            sht_origin.Select
            Cells(i, j + 1).Select
            Selection.Copy
            Set copy_time_rng = Selection
            
            If copy_time_rng = "" Then
                End
            End If
            
            
            Range((Cells(i, 1).Address)).Select
            Selection.Copy
            Set copy_word_rng = Selection
            tian = Day(copy_time_rng)   '为变量tian赋值
            yue = Month(copy_time_rng)  '为变量yue赋值
            
            
            ' 根据yue的值，进入特定的sheet
            Select Case yue
                Case 1
                    sht_jan.Select
                    ' 将粘贴板中的数据， 粘贴到当月当日对应的列的第一个空白单元格处
                    Cells((Cells(65535, tian).End(xlUp).Row + 1), tian).Select
                    ActiveSheet.Paste
                Case 2
                    sht_feb.Select
                    ' 将粘贴板中的数据， 粘贴到当月当日对应的列的第一个空白单元格处
                    Cells((Cells(65535, tian).End(xlUp).Row + 1), tian).Select
                    ActiveSheet.Paste
                Case 3
                    sht_march.Select
                    ' 将粘贴板中的数据， 粘贴到当月当日对应的列的第一个空白单元格处
                    Cells((Cells(65535, tian).End(xlUp).Row + 1), tian).Select
                    ActiveSheet.Paste
                Case 4
                    sht_april.Select
                    ' 将粘贴板中的数据， 粘贴到当月当日对应的列的第一个空白单元格处
                    Cells((Cells(65535, tian).End(xlUp).Row + 1), tian).Select
                    ActiveSheet.Paste
                Case 5
                    sht_may.Select
                    ' 将粘贴板中的数据， 粘贴到当月当日对应的列的第一个空白单元格处
                    Cells((Cells(65535, tian).End(xlUp).Row + 1), tian).Select
                    ActiveSheet.Paste
                Case 6
                    sht_june.Select
                    ' 将粘贴板中的数据， 粘贴到当月当日对应的列的第一个空白单元格处
                    Cells((Cells(65535, tian).End(xlUp).Row + 1), tian).Select
                    ActiveSheet.Paste
                Case 7
                    sht_july.Select
                    ' 将粘贴板中的数据， 粘贴到当月当日对应的列的第一个空白单元格处
                    Cells((Cells(65535, tian).End(xlUp).Row + 1), tian).Select
                    ActiveSheet.Paste
                Case 8
                    sht_aug.Select
                    ' 将粘贴板中的数据， 粘贴到当月当日对应的列的第一个空白单元格处
                    Cells((Cells(65535, tian).End(xlUp).Row + 1), tian).Select
                    ActiveSheet.Paste
                Case 9
                    sht_sep.Select
                    ' 将粘贴板中的数据， 粘贴到当月当日对应的列的第一个空白单元格处
                    Cells((Cells(65535, tian).End(xlUp).Row + 1), tian).Select
                    ActiveSheet.Paste
                Case 10
                    sht_oct.Select
                    ' 将粘贴板中的数据， 粘贴到当月当日对应的列的第一个空白单元格处
                    Cells((Cells(65535, tian).End(xlUp).Row + 1), tian).Select
                    ActiveSheet.Paste
                Case 11
                    sht_nov.Select
                    ' 将粘贴板中的数据， 粘贴到当月当日对应的列的第一个空白单元格处
                    Cells((Cells(65535, tian).End(xlUp).Row + 1), tian).Select
                    ActiveSheet.Paste
                Case 12
                    sht_dec.Select
                    ' 将粘贴板中的数据， 粘贴到当月当日对应的列的第一个空白单元格处
                    Cells((Cells(65535, tian).End(xlUp).Row + 1), tian).Select
                    ActiveSheet.Paste
                Case Else
                     MsgBox "Check whether there are some errors Please!"
            End Select
        Next j
    Next i
    
End Sub


