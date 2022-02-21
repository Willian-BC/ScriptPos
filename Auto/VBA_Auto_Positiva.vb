Sub atualizar()
    Application.ScreenUpdating = False
    For i = 2 To ThisWorkbook.Sheets.Count - 3
        Sheets(i).Activate
        If i = 2 Then
            aux = "atc"
            Cells(2, 7).FormulaR1C1 = "=COUNTIF('bd vrj'!C[-6],RC[-6])"
            Cells(2, 8).FormulaR1C1 = "=COUNTIF('bd blc'!C[-7],RC[-7])"
            Range("G2:H" & Cells(2, 1).End(xlDown).Row).Copy
            Range("G2:H" & Cells(2, 1).End(xlDown).Row).PasteSpecial Paste:=xlPasteValues
        ElseIf i = 3 Then
            aux = "vrj"
            Cells(2, 7).FormulaR1C1 = "=COUNTIF('bd atc'!C[-6],RC[-6])"
            Cells(2, 8).FormulaR1C1 = "=COUNTIF('bd blc'!C[-7],RC[-7])"
            Range("G2:H" & Cells(2, 1).End(xlDown).Row).Copy
            Range("G2:H" & Cells(2, 1).End(xlDown).Row).PasteSpecial Paste:=xlPasteValues
        Else
            aux = "blc"
            Cells(2, 7).FormulaR1C1 = "=COUNTIF('bd atc'!C[-6],RC[-6])"
            Cells(2, 8).FormulaR1C1 = "=COUNTIF('bd vrj'!C[-7],RC[-7])"
            Range("G2:H" & Cells(2, 1).End(xlDown).Row).Copy
            Range("G2:H" & Cells(2, 1).End(xlDown).Row).PasteSpecial Paste:=xlPasteValues
        End If
        ActiveWorkbook.Worksheets("bd " & aux).ListObjects("Tabela_" & aux).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("bd " & aux).ListObjects("Tabela_" & aux).Sort.SortFields.Add Key:=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending
        ActiveWorkbook.Worksheets("bd " & aux).ListObjects("Tabela_" & aux).Sort.SortFields.Add Key:=Range("E:E"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "janeiro,fevereiro,março,abril,maio,junho,julho,agosto,setembro,outubro,novembro,dezembro"
        ActiveWorkbook.Worksheets("bd " & aux).ListObjects("Tabela_" & aux).Sort.SortFields.Add Key:=Range("D:D"), SortOn:=xlSortOnValues, Order:=xlAscending
        With ActiveWorkbook.Worksheets("bd " & aux).ListObjects("Tabela_" & aux).Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    Next
    For i = 2 To ThisWorkbook.Sheets.Count - 3
        Sheets(i).Activate
        If i = 2 Then
            Cells(2, 9).FormulaR1C1 = "=IF(RC[-2]>0,IF(VLOOKUP(RC[-8],'bd vrj'!C[-8]:C[-5],4,0)<RC[-5],""varejo"",""atacado""),IF(RC[-1]>0,IF(VLOOKUP(RC[-8],'bd blc'!C[-8]:C[-5],4,0)<RC[-5],""balcão"",""atacado""),""atacado""))"
        ElseIf i = 3 Then
            Cells(2, 9).FormulaR1C1 = "=IF(RC[-2]>0,IF(VLOOKUP(RC[-8],'bd atc'!C[-8]:C[-5],4,0)<RC[-5],""atacado"",""varejo""),IF(RC[-1]>0,IF(VLOOKUP(RC[-8],'bd blc'!C[-8]:C[-5],4,0)<RC[-5],""balcão"",""varejo""),""varejo""))"
        Else
            Cells(2, 9).FormulaR1C1 = "=IF(RC[-2]>0,IF(VLOOKUP(RC[-8],'bd atc'!C[-8]:C[-5],4,0)<RC[-5],""atacado"",""balcão""),IF(RC[-1]>0,IF(VLOOKUP(RC[-8],'bd vrj'!C[-8]:C[-5],4,0)<RC[-5],""varejo"",""balcão""),""balcão""))"
        End If
        Cells(2, 10).FormulaR1C1 = "=IF(RC[-9]=R[-1]C[-9],R[-1]C + 1, 1)"
        Cells(2, 11).FormulaR1C1 = "=IF(RC[-5]=2021,""NOVO"",IF(COUNTIF('Positivado 2020'!C[-10],RC[-10])>0,""POSITIVADO"",""REATIVADO""))"
        Range("I2:K" & Cells(2, 1).End(xlDown).Row).Copy
        Range("I2:K" & Cells(2, 1).End(xlDown).Row).PasteSpecial Paste:=xlPasteValues
    Next
    ActiveWorkbook.RefreshAll
    Sheets(2).Activate
    Application.ScreenUpdating = True
    
End Sub

Sub limpar()
    Application.ScreenUpdating = False
    For i = 2 To ThisWorkbook.Sheets.Count - 3
        Sheets(i).Activate
        If i = 2 Then
            ActiveSheet.ListObjects("Tabela_atc").ShowAutoFilter = True
        ElseIf i = 3 Then
            ActiveSheet.ListObjects("Tabela_vrj").ShowAutoFilter = True
        Else
            ActiveSheet.ListObjects("Tabela_blc").ShowAutoFilter = True
        End If
        On Error Resume Next
        ActiveSheet.ShowAllData
        Range("A2", Cells(2, 11).End(xlDown)).ClearContents
    Next
    Sheets(2).Activate
    Application.ScreenUpdating = True
End Sub
