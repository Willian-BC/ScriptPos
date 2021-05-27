Sub Jornada()
    
    Application.ScreenUpdating = False

    Sheets("Planilha1").Activate
    If Not ActiveSheet.AutoFilterMode Then
        Range("B3", Cells(3, 2).End(xlToRight)).AutoFilter
    End If
    On Error Resume Next
    ActiveSheet.ShowAllData
    tam = (Cells(4, 12).End(xlToRight).Column - Cells(4, 11).Column) / 2
    
    For k = 1 To 6
        j = 10
        For i = 1 To tam
            ActiveSheet.Range("M:M").AutoFilter Field:=j + 2, Criteria1:=k
            j = j + 2
        Next
        lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        If Not IsEmpty(Cells(lin_inicio, 2).Value) Then
            If k < 4 Then
                Cells(lin_inicio, 37).Value = "CONTANTE ABAIXO"
            ElseIf k = 4 Then
                Cells(lin_inicio, 37).Value = "CONTANTE ESPERADO"
            ElseIf k > 4 Then
                Cells(lin_inicio, 37).Value = "CONTANTE ACIMA"
            End If
            Range("AK" & lin_inicio & ":AK" & Cells(lin_inicio, 2).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
        On Error Resume Next
        ActiveSheet.ShowAllData
    Next
    
    ActiveSheet.Range("E:E").AutoFilter Field:=4, Criteria1:="*VENDEDOR*"
    If Not IsEmpty(Cells(lin_inicio, 2).Value) Then
        lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        Cells(lin_inicio, 37).Value = "INATIVO"
        Range("AK" & lin_inicio & ":AK" & Cells(lin_inicio, 2).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
    End If
    ActiveSheet.ShowAllData
    ActiveSheet.Range("E:E").AutoFilter Field:=4, Criteria1:="INDISPONÍVEL", Operator:=xlOr, Criteria2:="INATIVORRR"
    If Not IsEmpty(Cells(lin_inicio, 2).Value) Then
        lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        Cells(lin_inicio, 37).Value = "INATIVO"
        Range("AK" & lin_inicio & ":AK" & Cells(lin_inicio, 2).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
    End If
    ActiveSheet.ShowAllData
    
    Set rngAF = Range("B4:B" & Cells(4, 2).End(xlDown).Row)
    ActiveSheet.Range("AK:AK").AutoFilter Field:=36, Criteria1:="="
    
    For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
        lin = rngcell.Row
        j = 11
        a = 0
        b = 0
        c = 0
INICIO:
        For i = 1 To tam
            If i = 1 Then
                Do While Cells(lin, j + 2).Value = 0
                    j = j + 2
                    tam = tam - 1
                    aux = 1
                    aux2 = 1
                Loop
                If aux > 0 Then
                    aux = 0
                    GoTo INICIO
                End If
                ref1 = Cells(lin, j + 2).Value
            ElseIf i = tam Then
                ref2 = Cells(lin, j + 2).Value
            ElseIf Cells(lin, j + 2).Value = ref1 Then
                a = a + 1
            ElseIf Cells(lin, j + 2).Value < ref1 Then
                b = b + 1
            ElseIf Cells(lin, j + 2).Value > ref1 Then
                c = c + 1
            End If
            j = j + 2
        Next
        If tam < 3 Then
            Cells(lin, 37).Value = "AGUARDANDO DADOS"
        ElseIf ref1 = ref2 And b > 0 And c = 0 Then
            Cells(lin, 37).Value = "ATENÇÃO"
        ElseIf ref1 = ref2 And b = 0 And c > 0 Then
            Cells(lin, 37).Value = "POTENCIAL"
        ElseIf ref1 > ref2 And c = 0 Then
            Cells(lin, 37).Value = "DECLÍNIO"
        ElseIf ref1 < ref2 And b = 0 Then
            Cells(lin, 37).Value = "CRESCIMENTO"
        Else
            Cells(lin, 37).Value = "OSCILANTE"
        End If
        If aux2 > 0 Then
            aux2 = 0
            tam = (Cells(4, 12).End(xlToRight).Column - Cells(4, 11).Column) / 2
        End If
    Next rngcell
    ActiveSheet.ShowAllData
    Cells(3, 2).Select
    Application.ScreenUpdating = True
End Sub