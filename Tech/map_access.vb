Sub map()

Dim ws As Worksheet
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
ActiveSheet.DisplayPageBreaks = False
Application.Calculation = xlManual

Set ws = Sheets("Sheet4")
ws.ShowAllData
ws.Range("C2:F" & ws.Cells(2, 3).End(xlDown).Row).ClearContents
lin = 2
lin_inicio = 2
ws.Cells(1, 1).Select
Do While Not IsEmpty(ws.Cells(Selection.Row, 1).Row)
    
    ws.Columns("A:A").Find(What:="StrProcesso_MMT =", After:=ActiveCell _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Activate
    ws.Columns("A:A").FindNext(After:=ActiveCell).Activate
    ws.Columns("A:A").FindNext(After:=ActiveCell).Activate
    'If lin = 2 Then ws.Columns("A:A").FindNext(After:=ActiveCell).Activate
    
    If Selection.Row < x Then GoTo FIM
    
    larray = Split(Selection.Value, """")
    ws.Cells(lin, 3) = larray(1)
    
    x = Selection.Row
    ws.Columns("A:A").Find(What:="Else", After:=ActiveCell _
        , LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False).Activate
    y = Selection.Row
    
    For x = x + 2 To y - 1
        If InStr(Cells(x, 1), "DoCmd.OpenReport") Then
            larray = Split(Cells(x, 1), """")
            ws.Cells(lin, 4) = larray(1)
            lin = lin + 1
            x = x + 1
        End If
    Next
    
    ws.Range("C" & lin_inicio & ":C" & Cells(lin_inicio, 4).End(xlDown).Row).FillDown
    lin_inicio = lin
Loop
FIM:
Application.Calculation = xlAutomatic
Cells(2, 5).FormulaR1C1 = "=IF(RC[-2]=R[-1]C[-2],R[-1]C + 1, 1)"
Cells(2, 6).FormulaR1C1 = "=IF(IFERROR(VLOOKUP(RC[-2],'Descrição'!C[-3]:C[-2],2,0),"""")=0,"""",IFERROR(VLOOKUP(RC[-2],'Descrição'!C[-3]:C[-2],2,0),""""))"
Range("E2:F" & Cells(2, 3).End(xlDown).Row).FillDown
Columns("E:E").Copy
Columns("E:E").PasteSpecial Paste:=xlPasteValues
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
ActiveSheet.DisplayPageBreaks = True
End Sub
