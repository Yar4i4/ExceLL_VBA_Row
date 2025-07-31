Attribute VB_Name = "Row"


Sub СтрокуУд()
PS = Range("L" & Rows.Count).End(xlUp).row
For i = PS To 16 Step -1
      ' If Cells(i, 2) = "" Or Mid(Cells(i, 2), 1, 1) = "." Or Cells(i, 2).Value Like "*...*" Or Cells(i, 2).Value Like "*Вода*" Or Cells(i, 2).Value Like "*Транспорт*" Or Cells(i, 2).Value Like "*Надбавка*" Or Cells(i, 2).Value Like "*Доплата*" Or Mid(Cells(i, 2), 1, 1) = Chr(133) Then
If Cells(i, 12).Value Like "*вернуть*" Then
Rows(i).Delete
End If
Next
 ' закрыть Макрос
Workbooks("макрос.xlsm").Close SaveChanges:=False
End Sub

Sub СтрокуВыделитьДобавить()
Application.DisplayAlerts = False 'отключаем показ сообщений
PS = Range("L" & Rows.Count).End(xlUp).row
For i = PS To 16 Step -1
      ' If Cells(i, 2) = "" Or Mid(Cells(i, 2), 1, 1) = "." Or Cells(i, 2).Value Like "*...*" Or Cells(i, 2).Value Like "*Вода*" Or Cells(i, 2).Value Like "*Транспорт*" Or Cells(i, 2).Value Like "*Надбавка*" Or Cells(i, 2).Value Like "*Доплата*" Or Mid(Cells(i, 2), 1, 1) = Chr(133) Then
If Cells(i, 12).Value Like "в" Then
'If Cells(i, 11).Value Like "*вернуть*" Then
Rows(i + 1).Insert
Cells(i + 1, 12) = "досписать"
Rows(i).Interior.colorIndex = 15
Cells(i + 1, "A").Borders.Weight = xlThin
Cells(i + 1, "B").Borders.Weight = xlThin
Cells(i + 1, "C").Borders.Weight = xlThin
Cells(i + 1, "D").Borders.Weight = xlThin
Cells(i + 1, "E").Borders.Weight = xlThin
Cells(i + 1, "F").Borders.Weight = xlThin
Cells(i + 1, "G").Borders.Weight = xlThin
Cells(i + 1, "H").Borders.Weight = xlThin
Cells(i + 1, "I").Borders.Weight = xlThin
Cells(i + 1, "J").Borders.Weight = xlThin
Cells(i + 1, "K").Borders.Weight = xlThin
End If
Next
Range("L1:L2222").Replace "в", "вернуть", xlPart

Range("B1:B1111").Replace " ТУ 1380-123-05757848-2014 из стали марки С345 с дополнительными требованиями по ударной вязкости не менее 4 кгс.м/см2 при температуре минус 40 градусов по ГОСТ 19281-2014", "", xlPart
Range("B1:B1111").Replace " ТУ 1380-123-05757848-2014 из стали марки С345 с дополнительными требованиями по ударной вязкости не менее 4 кгс.м/см2 при температуре минус 40 градусов по ГОСТ 19281-2014", "", xlPart




 ' закрыть Макрос
Workbooks("макрос.xlsm").Close SaveChanges:=False
Application.DisplayAlerts = True 'обязательно вкл. показ сообщений
End Sub




Sub СтрокуНовЛист()
Dim i&
Application.ScreenUpdating = False
For i = 11 To Cells(Rows.Count, 2).End(xlUp).row
'If Cells(i, 2) = "Паспорт" Or Cells(i, 6) = "Монитор" Or Cells(i, 6) = "Жесткий диск" Then
If Cells(i, 2).Value Like "пасп*" Then
Rows(i).Copy
Sheets("Лист2").Cells(Sheets("Лист2").Cells(Rows.Count, 2).End(xlUp).row + 1, 1).PasteSpecial
End If
Next
Application.ScreenUpdating = True
Application.CutCopyMode = False
End Sub

Sub УдалиЧётСтроку()
Attribute УдалиЧётСтроку.VB_ProcData.VB_Invoke_Func = " \n14"

Dim i As Integer
For i = 10000 To 1 Step -1 'вместо 10 000 последнюю строку Вашей таблицы.
If i Mod 2 = 0 Then '0 - Для четных. Для нечетных 1
Rows(i).Delete
End If
Next
    ' закрыть Макрос
Workbooks("макрос.xlsm").Close SaveChanges:=False
End Sub
Sub Вставить123_P_AAзпт()
    'выделить фон, если пустой объём
    r0_ = 4
    r1_ = Range("AA" & Rows.Count).End(xlUp).row
    For U = r0_ To r1_
    If Range("AA" & U) = "," Then
'    Range("D" & U).Font.Color = vbRed
    Range("P" & U).Copy
    Range("P" & U).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    End If
    Next U
    End Sub
    ' закрыть Макрос
Workbooks("макрос.xlsm").Close SaveChanges:=False
End Sub


Sub ВставитьДвеСтроки()
    Dim lLastRow As Long, li As Long
    Application.ScreenUpdating = 0
    lLastRow = Cells(Rows.Count, 1).End(xlUp).row
    For li = lLastRow To 1 Step -1
        Rows(li).Resize(2).Insert
    Next li
    Application.ScreenUpdating = 1
End Sub
Sub ВставитьТриСтроки()
    Dim lLastRow As Long, li As Long
    Application.ScreenUpdating = 0
    lLastRow = Cells(Rows.Count, 1).End(xlUp).row
    For li = lLastRow To 1 Step -1
        Rows(li).Resize(3).Insert
    Next li
    Application.ScreenUpdating = 1
End Sub

Sub УдалитьПустыеСтроки()
          PS = Range("B" & Rows.Count).End(xlUp).row
       For i = PS To 1 Step -1
       If Cells(i, 1) = "" Or Mid(Cells(i, 1), 1, 1) = "." Or Mid(Cells(i, 1), 1, 1) = Chr(133) Then
         Rows(i).Delete
        End If
        Next
End Sub




