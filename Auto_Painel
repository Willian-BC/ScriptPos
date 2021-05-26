Sub prospecção_carteira()
    Application.ScreenUpdating = False
    
    On Error Resume Next
    Sheets("CAMPANHA").Activate
    Columns("A:A").Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    Cells(1, 1).Value = Cells(1, 2).Value
    Cells(2, 1).FormulaR1C1 = "=VALUE(RC[1])"
    Cells(2, 17).Formula2R1C1 = "=SUM(VALUE(RC[-4]:RC[-1]))"
    Range("A2:A" & Cells(2, 2).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
    Range("Q2:Q" & Cells(2, 2).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
    Columns("A:Q").Copy
    Columns("A:Q").PasteSpecial Paste:=xlPasteValues
    Columns("B:B").Delete
    Sheets("CARTEIRA").Activate
    If Not ActiveSheet.AutoFilterMode Then
        Range("A1").End(xlToRight).AutoFilter
    End If
    On Error Resume Next
    ActiveSheet.ShowAllData
    ActiveSheet.Range("G:G").AutoFilter Field:=7, Criteria1:="VENDEDOR VAREJO OFICINA"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    Rows(lin_inicio & ":" & lin_fim).SpecialCells(xlCellTypeVisible).Delete
    ActiveSheet.ShowAllData
    Cells(1, 8).Value = "TEMPO NA CARTEIRA"
    Cells(1, 9).Value = "CÓD. CNAE"
    Cells(1, 10).Value = "DESCRIÇÃO CNAE"
    Cells(1, 11).Value = "ORIGEM"
    Cells(1, 12).Value = "CONTATO ÚLTIMOS 30 DIAS"
    Cells(1, 13).Value = "COMPRA"
    Cells(1, 14).Value = "FIDELIZAÇÃO"
    mes = Month(Date)
    j = 6
    For i = 1 To mes
        Cells(1, i + 14).Value = MonthName(i)
        Cells(2, i + 14).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-" & i + 13 & "],'FAT. GERAL'!C[-" & i + 13 & "]:C[-" & mes - j & "]," & i + 1 & ",0),0)"
        j = j - 1
    Next
    Cells(1, i + 14).Value = "PROJEÇÃO " & UCase(MonthName(mes))
    Cells(1, i + 15).Value = "RECOMPRA"
    Cells(1, i + 16).Value = "TOTAL FATURAMENTO"
    Cells(1, i + 17).Value = "MÉDIA CLIENTE"
    Cells(2, 9).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-7],CNAE!C[-8]:C[-7],2,0),""SEM CNAE"")"
    Cells(2, 10).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-8],CNAE!C[-9]:C[-7],3,0),""SEM CNAE"")"
    Cells(2, 11).Value = "=IFERROR(VLOOKUP(RC[-10],ORIGEM!C[-10]:C[-9],2,0),"""")"
    Cells(2, 12).FormulaR1C1 = _
    "=IF(SUM(IFERROR(VLOOKUP(RC[-11],CHAMADAS!C[-11]:C[-10],2,0),0)+IFERROR(VLOOKUP(RC[-11],'QTD. VENDAS'!C[-11]:C[-10],2,0),0)+IFERROR(VLOOKUP(RC[-11],'QTD. COTAÇÕES'!C[-11]:C[-10],2,0),0)+IFERROR(VLOOKUP(RC[-11],CAMPANHA!C[-11]:C[4],16,0),0))>0,""CONTATO"",""SEM CONTATO"")"
    Cells(2, 13).FormulaR1C1 = "=IF(SUM(RC[2]:RC[" & mes + 1 & "])>0,""COMPRANTE"",""NÃO COMPRANTE"")"
    Cells(2, 14).FormulaR1C1 = "=IF(RC[" & mes + 2 & "]=0,""NÃO PROSPECTADO"",IF(AND(RC[" & mes + 2 & "]>=ROUNDUP(" & mes & "/2,0),AVERAGE(RC[1]:RC[" & mes & "])>VLOOKUP(RC[-8],'MÉDIA'!C[-13]:C[-12],2,0)),""FIDELIZADO"",""EM DESENVOLVIMENTO""))"
    'Cells(2, i + 14).FormulaR1C1 = "PROJEÇÃO"
    Cells(2, i + 15).FormulaR1C1 = "=COUNTIF(RC[-" & mes + 1 & "]:RC[-2],"">0"")"
    Cells(2, i + 16).FormulaR1C1 = "=SUM(RC[-" & mes + 2 & "]:RC[-3])"
    Cells(2, i + 17).FormulaR1C1 = "=AVERAGE(RC[-" & mes + 3 & "]:RC[-4])"
    
    col = Split(Cells(1, 1).End(xlToRight).Address, "$")(1)
    Range("I2:" & col & Cells(2, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
    Columns("I:" & col).Copy
    Columns("I:" & col).PasteSpecial Paste:=xlPasteValues
    
    Application.ScreenUpdating = True
End Sub

Sub prospecção_positivação()
    Application.ScreenUpdating = False
    Sheets("POSITIVAÇÃO ATC").Range("A1").End(xlToRight).Copy Sheets("BASE POSITIVAÇÃO").Cells(1, 1)
    For i = 1 To 2
        If i = 1 Then
            aux = "ATC"
        Else
            aux = "VRJ"
        End If
        Sheets("POSITIVAÇÃO " & aux).Activate
        ActiveWorkbook.Worksheets("POSITIVAÇÃO " & aux).Sort.SortFields.Clear
        ActiveWorkbook.Worksheets("POSITIVAÇÃO " & aux).Sort.SortFields.Add2 Key:=Range("A:A"), SortOn:=xlSortOnValues, Order:=xlAscending
        ActiveWorkbook.Worksheets("POSITIVAÇÃO " & aux).Sort.SortFields.Add2 Key:=Range("E:E"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
            "janeiro,fevereiro,março,abril,maio,junho,julho,agosto,setembro,outubro,novembro,dezembro"
        ActiveWorkbook.Worksheets("POSITIVAÇÃO " & aux).Sort.SortFields.Add2 Key:=Range("D:D"), SortOn:=xlSortOnValues, Order:=xlAscending
        With ActiveWorkbook.Worksheets("POSITIVAÇÃO " & aux).Sort
            .SetRange Range("A:E")
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
        Cells(2, 6).FormulaR1C1 = "=IF(RC[-5]=R[-1]C[-5],0,1)"
        Range("F2:F" & Cells(2, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        Columns("F:F").Copy
        Columns("F:F").PasteSpecial Paste:=xlPasteValues
        If Not ActiveSheet.AutoFilterMode Then
            Range("A1").End(xlToRight).AutoFilter
        End If
        On Error Resume Next
        ActiveSheet.ShowAllData
        ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:=1
        If i = 1 Then
            ActiveSheet.Range("B:B").AutoFilter Field:=2, Criteria1:=Array("EQUIPE - PROSPECÇÃO A", "EQUIPE - PROSPECÇÃO B", "EQUIPE - PROSPECÇÃO C"), Operator:=xlFilterValues
            lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            lin_fim = Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
            Range("A" & lin_inicio & ":E" & lin_fim).Copy Sheets("BASE POSITIVAÇÃO").Cells(2, 1)
        Else
            ActiveSheet.Range("B:B").AutoFilter Field:=2, Criteria1:=Array("EQUIPE PROSPECÇÃO OFICINA", "EQUIPE PROSPECÇÃO OFICINA 2", "EQUIPE PROSPECÇÃO OFICINA 3 (FROTAS)"), Operator:=xlFilterValues
            lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            lin_fim = Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
            Range("A" & lin_inicio & ":E" & lin_fim).Copy Sheets("BASE POSITIVAÇÃO").Cells(Cells(2, 1).End(xlDown).Row, 1)
        End If
        
    Next
    Sheets("BASE POSITIVAÇÃO").Activate
    ActiveWorkbook.Worksheets("BASE POSITIVAÇÃO").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("BASE POSITIVAÇÃO").Sort.SortFields.Add2 Key:=Range("B:B"), SortOn:=xlSortOnValues, Order:=xlAscending
    ActiveWorkbook.Worksheets("BASE POSITIVAÇÃO").Sort.SortFields.Add2 Key:=Range("E:E"), SortOn:=xlSortOnValues, Order:=xlAscending, CustomOrder:= _
        "janeiro,fevereiro,março,abril,maio,junho,julho,agosto,setembro,outubro,novembro,dezembro"
    ActiveWorkbook.Worksheets("BASE POSITIVAÇÃO").Sort.SortFields.Add2 Key:=Range("D:D"), SortOn:=xlSortOnValues, Order:=xlAscending
    With ActiveWorkbook.Worksheets("BASE POSITIVAÇÃO").Sort
        .SetRange Range("A:E")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Cells(2, 6).FormulaR1C1 = "=IF(RC[-1]=R[-1]C[-1],R[-1]C+1,1)"
    Range("F2:F" & Cells(2, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
    Columns("F:F").Copy
    Columns("F:F").PasteSpecial Paste:=xlPasteValues

    Application.ScreenUpdating = True
    
End Sub

Sub prospecção_apagar()
    For i = 1 To Sheets.Count
        Sheets(i).Activate
        On Error Resume Next
        ActiveSheet.ShowAllData
        Cells.ClearContents
    Next
    Sheets(1).Activate
End Sub
