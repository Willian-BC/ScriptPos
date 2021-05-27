Sub WORKLOAD()

Dim nome As String
Dim lin1 As Byte, lin2 As Byte
Dim soma1 As Double, soma2 As Double, hora As Double

Application.ScreenUpdating = False

Sheets("WORKLOAD").Activate
Sheets("WORKLOAD").Range("B3:C14").ClearContents
Sheets("WORKLOAD").Range("B47:C58").ClearContents

lin1 = 3
lin2 = 47
Do While Not IsEmpty(Range("A" & lin1))
    nome = Range("A" & lin1).Value
    Sheets("PLANILHA CONTROLE PROJETOS").Select
    
    'ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=7, Criteria1:="=*" & nome & "*"
    ActiveSheet.ListObjects("Table1").ShowAutoFilter = True
    On Error Resume Next
    ActiveSheet.ShowAllData
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=7, Criteria1:="=*" & nome & "*"
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=17, Criteria1:="PENDENTE", Operator:=xlOr, Criteria2:="ATRASADO"
    soma2 = WorksheetFunction.Subtotal(109, Range("X:X"))
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=11, Criteria1:="<" & Format((Date + 90), "mm/dd/yyyy")
    soma1 = WorksheetFunction.Subtotal(109, Range("U:U"))
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=11
    
    Sheets("WORKLOAD").Select
    Range("B" & lin1).Value = (soma1 * 8.5)
    Range("B" & lin2).Value = (soma2 * 8.5)
    
    Sheets("PLANILHA CONTROLE PROJETOS").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=7
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=8, Criteria1:="=*" & nome & "*"
    soma2 = WorksheetFunction.Subtotal(109, Range("X:X"))
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=11, Criteria1:="<" & Format((Date + 90), "mm/dd/yyyy")
    soma1 = WorksheetFunction.Subtotal(109, Range("U:U"))
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=11
    
    Sheets("WORKLOAD").Select
    Range("C" & lin1).Value = (soma1 * 8.5)
    Range("C" & lin2).Value = (soma2 * 8.5)
    
    lin1 = lin1 + 1
    lin2 = lin2 + 1
Loop

Sheets("PLANILHA CONTROLE PROJETOS").Select
On Error Resume Next
ActiveSheet.ShowAllData

Call WORKLOAD_CELULA

Sheets("WORKLOAD").Select

Sheets("DADOS GRÁFICOS").Select
ActiveWorkbook.RefreshAll
Sheets("DADOS TABELAS").Select
ActiveWorkbook.RefreshAll

Sheets("WORKLOAD").Select

Application.ScreenUpdating = True

End Sub

Sub WORKTIME()

Dim nome As String
Dim lin As Byte

Application.ScreenUpdating = False

i1 = 1
Sheets("WORKTIME").Activate

While i1 < 1500
    Do While IsEmpty(Range("A" & i1))
        If i1 < 1500 Then
            i1 = i1 + 1
        Else
            Exit Do
        End If
    Loop
    i1 = i1 + 1
    Do While Not IsEmpty(Range("A" & i1))
        Range("A" & i1).CurrentRegion.Select
        Selection.ClearContents
        Range("A" & i1).Value = "NOME DO PROJETO"
        Range("B" & i1).Value = "DATA DE INICIO"
        Range("C" & i1).Value = "DIAS PROJETO"
        Range("C" & i1 + 1).FormulaR1C1 = "=RC[1]-RC[-1]"
        Range("D" & i1).Value = "DATA PRAZO"
        i1 = i1 + 1
    Loop
Wend

Sheets("PLANILHA CONTROLE PROJETOS").Select
On Error Resume Next
ActiveSheet.ShowAllData
Range("B17:AA" & Cells(Rows.Count, 8).End(xlUp).Row).Replace What:="", Replacement:="NULL", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Sheets("WORKLOAD").Activate

lin = 3
a = 1
b = 1
c = 1
d = 1
Do While Not IsEmpty(Range("A" & lin))
    nome = Range("A" & lin).Value

    Sheets("PLANILHA CONTROLE PROJETOS").Select
    If ActiveSheet.AutoFilterMode = 0 Then
        Range("B17", Selection.End(xlToRight)).AutoFilter
    End If
    On Error Resume Next
    ActiveSheet.ShowAllData
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=7, Criteria1:="=*" & nome & "*"

INICIO:
    If Application.WorksheetFunction.Subtotal(103, Columns("B")) > 1 Then
        Range("C17:C" & Cells(Rows.Count, 3).End(xlUp).Row).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
        Sheets("WORKTIME").Select
        Range("A" & a).End(xlDown).Select
        a = ActiveCell.Row
        a1 = a + 1
        Range("A" & a - 1).Value = nome
        Range("A" & a).Select
        ActiveCell.PasteSpecial Paste:=xlPasteValues
        Range("A" & a).End(xlDown).Select
        a = ActiveCell.Row

        Sheets("PLANILHA CONTROLE PROJETOS").Select
        Range("L17:L" & Cells(Rows.Count, 3).End(xlUp).Row).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
        Sheets("WORKTIME").Select
        Range("B" & b).End(xlDown).Select
        b = ActiveCell.Row
        Range("B" & b).Select
        ActiveCell.PasteSpecial Paste:=xlPasteValues
        Range("B" & b).End(xlDown).Select
        b = ActiveCell.Row

        Sheets("PLANILHA CONTROLE PROJETOS").Select
        Range("M17:M" & Cells(Rows.Count, 3).End(xlUp).Row).Select
        Selection.SpecialCells(xlCellTypeVisible).Select
        Selection.Copy
        Sheets("WORKTIME").Select
        Range("D" & d).End(xlDown).Select
        d = ActiveCell.Row
        Range("D" & d).Select
        ActiveCell.PasteSpecial Paste:=xlPasteValues
        Range("D" & d).End(xlDown).Select
        d = ActiveCell.Row

        Range("C" & c).End(xlDown).Select
        c = ActiveCell.Row
        Range("C" & c + 1).Select
        Selection.AutoFill Destination:=Range("C" & (c + 1) & ":C" & d)
        Range("C" & c).End(xlDown).Select
        c = ActiveCell.Row

        Worksheets("WORKTIME").ChartObjects("Chart " & (lin - 2)).Activate
        ActiveChart.SeriesCollection(1).XValues = Range("A" & a1 & ":A" & a)                  'eixo nome dos projetos (horizontal)
        ActiveChart.SeriesCollection(1).Values = Range("B" & a1 & ":B" & a)                       'eixo data de inicio
        ActiveChart.SeriesCollection(2).Values = Range("C" & a1 & ":C" & a)                       'eixo dias projeto


    Else
        ActiveSheet.ShowAllData
        ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=8, Criteria1:="=*" & nome & "*"
        GoTo INICIO:
    End If

    Sheets("WORKLOAD").Activate
    lin = lin + 1
Loop

Sheets("PLANILHA CONTROLE PROJETOS").Select
On Error Resume Next
ActiveSheet.ShowAllData
Range("B17:AA" & Cells(Rows.Count, 4).End(xlUp).Row).Replace What:="NULL", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Sheets("WORKTIME").Activate
Range("A1:D" & d).Replace What:="NULL", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False

Application.ScreenUpdating = True

End Sub

Sub DADOS_GRAFICO()
    
    Application.ScreenUpdating = False
    
    Dim PRange As Range
    Dim nome As String
    Dim lin As Byte
    
    Set PRange = Sheets(1).Range("C18").CurrentRegion
    
    Sheets("DADOS GRÁFICOS").PivotTables("PivotTable1").ChangePivotCache _
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange, Version:=6)
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    
 
    
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        Range("c18").CurrentRegion, Version:=6).CreatePivotTable _
        TableDestination:="DADOS GRÁFICOS!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=6
    Sheets(3).Select
    
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
    Selection).CreatePivotTable TableDestination:= _
    "DADOS GRÁFICOS!R3C1", TableName:="Pivot Table 1"
    

    
    PivotTable.RefreshTable
    
    Sheets("DADOS GRÁFICOS").Select
    Range("B3").CurrentRegion.Select
    Selection.ClearContents
    Range("C2").Value = "TEMPO ESTIMADO DE EXCEUÇÃO DO PROJETO"
    Range("D2").Value = "% CONSUMIDO DO TEMPO DISPONÍVEL PMB (HH)"
    
    Sheets("PLANILHA CONTROLE PROJETOS").Select
    If ActiveSheet.AutoFilterMode = 0 Then
        Range("B17", Selection.End(xlToRight)).AutoFilter
    End If
    On Error Resume Next
    ActiveSheet.ShowAllData
    
    Range("C17:C" & Cells(Rows.Count, 3).End(xlUp).Row).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Sheets("DADOS GRÁFICOS").Select
    Range("B2").Select
    ActiveSheet.Paste
    
    Sheets("PLANILHA CONTROLE PROJETOS").Select
    Range("F17:F" & Cells(Rows.Count, 3).End(xlUp).Row).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Sheets("DADOS GRÁFICOS").Select
    Range("E2").Select
    ActiveSheet.Paste
    
    Sheets("PLANILHA CONTROLE PROJETOS").Select
    Range("J17:J" & Cells(Rows.Count, 3).End(xlUp).Row).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    Sheets("DADOS GRÁFICOS").Select
    Range("F2").Select
    ActiveSheet.Paste
        
    Range("B:E").RemoveDuplicates Columns:=1, Header:=xlYes
    
    lin = 3
    Do While Not IsEmpty(Range("B" & lin))
        nome = Range("B" & lin).Value
        Sheets("PLANILHA CONTROLE PROJETOS").Select
        On Error Resume Next
        ActiveSheet.ShowAllData
        Range("C:C").AutoFilter Field:=2, Criteria1:="=*" & nome & "*"
        
        soma = WorksheetFunction.Subtotal(109, Range("U:U"))
        hora = (soma * 8.5)
        
        Sheets("DADOS GRÁFICOS").Select
        Range("C" & lin).Value = hora
        lin = lin + 1
        hora = 0
    Loop
    
    Sheets("PLANILHA CONTROLE PROJETOS").Select
    On Error Resume Next
    ActiveSheet.ShowAllData
    
    Sheets("DADOS GRÁFICOS").Select
    Range("D3").FormulaR1C1 = "=RC[-1]/SUM(WORKLOAD!R3C5:R14C5)"
    Range("D3").AutoFill Destination:=Range("D3:D" & Range("B3").End(xlDown).Row)
    ActiveSheet.PivotTables("PivotTable1").ChangePivotCache ActiveWorkbook. _
        PivotCaches.Create(SourceType:=xlDatabase, SourceData:=Range("B2").CurrentRegion, Version:=6)
    ActiveSheet.PivotTables("PivotTable1").PivotCache.Refresh
    'WorkSheets("DADOS GRÁFICOS").ChartObjects("Chart1").Activate
    
    Application.ScreenUpdating = True
    
End Sub

Sub WORKLOAD_CELULA()

Dim nome As String
Dim lin As Byte
Dim soma1 As Double, soma2 As Double, hora As Double

Sheets("WORKLOAD").Activate
Sheets("WORKLOAD").Range("B89:D95").ClearContents
Sheets("WORKLOAD").Range("R89:T95").ClearContents
Sheets("WORKLOAD").Range("B101:D104").ClearContents
Sheets("WORKLOAD").Range("R101:T104").ClearContents
Sheets("WORKLOAD").Range("B113:D115").ClearContents
Sheets("WORKLOAD").Range("R113:T115").ClearContents
Sheets("WORKLOAD").Range("B125:D126").ClearContents
Sheets("WORKLOAD").Range("R125:T126").ClearContents

lin = 89
aux = 1
INICIO:
Do While Not IsEmpty(Range("A" & lin))
    nome = Range("A" & lin).Value
    Sheets("PLANILHA CONTROLE PROJETOS").Select
    On Error Resume Next
    ActiveSheet.ShowAllData
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=7, Criteria1:="=*" & nome & "*"
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=17, Criteria1:="PENDENTE", Operator:=xlOr, Criteria2:="ATRASADO"
    If aux = 1 Then
        ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5, Criteria1:=Array("ARMAGEM", "ESPIRALAGEM", "EXTRUSÃO", "FITAS", "PERFILAGEM", _
            "PROCESSOS DE FABRICAÇÃO"), Operator:=xlFilterValues
    ElseIf aux = 2 Then
        ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5, Criteria1:=Array("NDT", "SOLDA"), Operator:=xlFilterValues
    ElseIf aux = 3 Then
        ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5, Criteria1:=Array("AMARRAÇÃO", "MMT"), Operator:=xlFilterValues
    ElseIf aux = 4 Then
        ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=5, Criteria1:=Array("CAPABILIDADE"), Operator:=xlFilterValues
    End If
    
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=4, Criteria1:="PROJETO"
    soma1 = WorksheetFunction.Subtotal(109, Range("X:X"))
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=11, Criteria1:="<" & Format((Date + 90), "mm/dd/yyyy")
    soma2 = WorksheetFunction.Subtotal(109, Range("U:U"))
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=11
    Sheets("WORKLOAD").Select
    Range("B" & lin).Value = (soma1 * 8.5)
    Range("R" & lin).Value = (soma2 * 8.5)
    
    Sheets("PLANILHA CONTROLE PROJETOS").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=4, Criteria1:="ATENDIMENTO"
    soma1 = WorksheetFunction.Subtotal(109, Range("X:X"))
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=11, Criteria1:="<" & Format((Date + 90), "mm/dd/yyyy")
    soma2 = WorksheetFunction.Subtotal(109, Range("U:U"))
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=11
    Sheets("WORKLOAD").Select
    Range("C" & lin).Value = (soma1 * 8.5)
    Range("S" & lin).Value = (soma2 * 8.5)
    
    Sheets("PLANILHA CONTROLE PROJETOS").Select
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=4, Criteria1:="ATIVIDADE"
    soma1 = WorksheetFunction.Subtotal(109, Range("X:X"))
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=11, Criteria1:="<" & Format((Date + 90), "mm/dd/yyyy")
    soma2 = WorksheetFunction.Subtotal(109, Range("U:U"))
    ActiveSheet.ListObjects("Table1").Range.AutoFilter Field:=11
    Sheets("WORKLOAD").Select
    Range("D" & lin).Value = (soma1 * 8.5)
    Range("T" & lin).Value = (soma2 * 8.5)
    
    lin = lin + 1
Loop
If aux = 1 Then
    lin = 101
    aux = aux + 1
    GoTo INICIO
ElseIf aux = 2 Then
    lin = 113
    aux = aux + 1
    GoTo INICIO
ElseIf aux = 3 Then
    lin = 125
    aux = aux + 1
    GoTo INICIO
End If

Sheets("PLANILHA CONTROLE PROJETOS").Select
On Error Resume Next
ActiveSheet.ShowAllData


End Sub
