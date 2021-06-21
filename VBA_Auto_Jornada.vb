Sub Jornada()
    
    Application.ScreenUpdating = False

    Sheets("Jornada").Activate
    If Not ActiveSheet.AutoFilterMode Then
        Range("A3", Cells(3, 2).End(xlToRight)).AutoFilter
    End If
    On Error Resume Next
    ActiveSheet.ShowAllData
    Range("AK4", Cells(4, 37).End(xlDown)).ClearContents
    tam = (Cells(4, 12).End(xlToRight).Column - Cells(4, 11).Column) / 2
    
    For k = 1 To 6
        j = 11
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
    
    ActiveSheet.Range("E:E").AutoFilter Field:=5, Criteria1:="*VENDEDOR*", Operator:=xlOr, Criteria2:="*CLIENTES*"
    If Not IsEmpty(Cells(lin_inicio, 2).Value) Then
        lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        Cells(lin_inicio, 37).Value = "INATIVO"
        Range("AK" & lin_inicio & ":AK" & Cells(lin_inicio, 2).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
    End If
    ActiveSheet.ShowAllData
    ActiveSheet.Range("E:E").AutoFilter Field:=5, Criteria1:="INDISPONÍVEL", Operator:=xlOr, Criteria2:="INATIVORRR"
    If Not IsEmpty(Cells(lin_inicio, 2).Value) Then
        lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        Cells(lin_inicio, 37).Value = "INATIVO"
        Range("AK" & lin_inicio & ":AK" & Cells(lin_inicio, 2).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
    End If
    ActiveSheet.ShowAllData
    
    Set rngAF = Range("B4:B" & Cells(4, 2).End(xlDown).Row)
    ActiveSheet.Range("AK:AK").AutoFilter Field:=37, Criteria1:="="
    
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
        ElseIf ref1 = ref2 And ref1 < 4 And a > 0 And b = 0 And c = 0 Then
            Cells(lin, 37).Value = "CONTANTE ABAIXO"
        ElseIf ref1 = ref2 And ref1 = 4 And a > 0 And b = 0 And c = 0 Then
            Cells(lin, 37).Value = "CONTANTE ESPERADO"
        ElseIf ref1 = ref2 And ref1 > 4 And a > 0 And b = 0 And c = 0 Then
            Cells(lin, 37).Value = "CONTANTE ACIMA"
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

Sub Classificar()
    Application.ScreenUpdating = False
    
    Sheets("Classificação").Activate
    If Not ActiveSheet.AutoFilterMode Then
        Range("A1", Cells(1, 1).End(xlToRight)).AutoFilter
    End If
    On Error Resume Next
    ActiveSheet.ShowAllData
    Cells(2, 9).FormulaR1C1 = "=RC[-1]/4"
    Range("I2:I" & Cells(2, 1).End(xlDown).Row).FillDown
    Columns("I:I").Copy
    Columns("I:I").PasteSpecial Paste:=xlPasteValues
    
'prospecção
    ActiveSheet.Range("B:B").AutoFilter Field:=2, Criteria1:="*PROSPECÇÃO*"
    ActiveSheet.Range("H:H").AutoFilter Field:=8, Criteria1:="<=0"
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="<>CONTATADO"
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="<>CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "RECUPERAÇÃO 1"
        Cells(lin_inicio, 11).Value = 1
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "PROSPECÇÃO"
        Cells(lin_inicio, 11).Value = 2
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    ActiveSheet.Range("F:F").AutoFilter Field:=6
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "PROSPECÇÃO"
        Cells(lin_inicio, 11).Value = 2
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    
    ActiveSheet.Range("H:H").AutoFilter Field:=8, Criteria1:=">0"
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="<>CONTATADO"
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="<>CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "RECUPERAÇÃO 2"
        Cells(lin_inicio, 11).Value = 3
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    
    ActiveSheet.Range("H:H").AutoFilter Field:=8
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="CONTATADO"
    ActiveSheet.Range("B:B").AutoFilter Field:=2, Criteria1:=Array( _
        "EQUIPE - PROSPECÇÃO A", "EQUIPE - PROSPECÇÃO B", "EQUIPE - PROSPECÇÃO C"), Operator:=xlFilterValues
    ActiveSheet.Range("I:I").AutoFilter Field:=9, Criteria1:=">" & Replace(Sheets("Ticket").Cells(18, 2).Value, ",", ".")
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "FIDELIZAÇÃO"
        Cells(lin_inicio, 11).Value = 5
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    ActiveSheet.Range("F:F").AutoFilter Field:=6
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "FIDELIZAÇÃO"
        Cells(lin_inicio, 11).Value = 5
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="<>CONTATADO"
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="CONTATADO"
    ActiveSheet.Range("B:B").AutoFilter Field:=2, Criteria1:=Array( _
        "EQUIPE PROSPECÇÃO OFICINA", "EQUIPE PROSPECÇÃO OFICINA 2", "EQUIPE PROSPECÇÃO OFICINA 3 (FROTAS)"), Operator:=xlFilterValues
    ActiveSheet.Range("I:I").AutoFilter Field:=9, Criteria1:=">" & Replace(Sheets("Ticket").Cells(19, 2).Value, ",", ".")
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "FIDELIZAÇÃO"
        Cells(lin_inicio, 11).Value = 5
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    ActiveSheet.Range("F:F").AutoFilter Field:=6
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "FIDELIZAÇÃO"
        Cells(lin_inicio, 11).Value = 5
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    
    ActiveSheet.Range("B:B").AutoFilter Field:=2, Criteria1:="*PROSPECÇÃO*"
    ActiveSheet.Range("I:I").AutoFilter Field:=9
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="<>CONTATADO"
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="CONTATADO"
    ActiveSheet.Range("J:J").AutoFilter Field:=10, Criteria1:="="
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "DESENVOLVIMENTO"
        Cells(lin_inicio, 11).Value = 4
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    ActiveSheet.Range("F:F").AutoFilter Field:=6
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "DESENVOLVIMENTO"
        Cells(lin_inicio, 11).Value = 4
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    On Error Resume Next
    ActiveSheet.ShowAllData

'relacionamento
    ActiveSheet.Range("B:B").AutoFilter Field:=2, Criteria1:="<>*PROSPECÇÃO*"
    ActiveSheet.Range("H:H").AutoFilter Field:=8, Criteria1:="<=0"
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="<>CONTATADO"
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="<>CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "PERDIDO"
        Cells(lin_inicio, 11).Value = 1
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "RECUPERAÇÃO 3"
        Cells(lin_inicio, 11).Value = 2
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    ActiveSheet.Range("F:F").AutoFilter Field:=6
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "RECUPERAÇÃO 3"
        Cells(lin_inicio, 11).Value = 2
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    
    ActiveSheet.Range("H:H").AutoFilter Field:=8, Criteria1:=">0"
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="<>CONTATADO"
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="<>CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).Value = "OBSERVAÇÃO"
        Cells(lin_inicio, 11).Value = 3
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="CONTATADO"
    ActiveSheet.Range("B:B").AutoFilter Field:=2, Criteria1:="*EQUIPE OFICINA*"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).FormulaR1C1 = "=IF(RC[-1]>(Ticket!R20C2*2),""REENQUADRAMENTO"",IF(RC[-1]>Ticket!R20C2,""RETENÇÃO"",""RELACIONAMENTO""))"
        Cells(lin_inicio, 11).FormulaR1C1 = "=IF(RC[-2]>(Ticket!R20C2*2),6,IF(RC[-2]>Ticket!R20C2,5,4))"
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    ActiveSheet.Range("F:F").AutoFilter Field:=6
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).FormulaR1C1 = "=IF(RC[-1]>(Ticket!R20C2*2),""REENQUADRAMENTO"",IF(RC[-1]>Ticket!R20C2,""RETENÇÃO"",""RELACIONAMENTO""))"
        Cells(lin_inicio, 11).FormulaR1C1 = "=IF(RC[-2]>(Ticket!R20C2*2),6,IF(RC[-2]>Ticket!R20C2,5,4))"
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="<>CONTATADO"
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="CONTATADO"
    ActiveSheet.Range("B:B").AutoFilter Field:=2, Criteria1:="EQUIPE ATACADO NOVO 1"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).FormulaR1C1 = "=IF(RC[-1]>(Ticket!R22C2),""REENQUADRAMENTO"",IF(RC[-1]>Ticket!R21C2,""RETENÇÃO"",""RELACIONAMENTO""))"
        Cells(lin_inicio, 11).FormulaR1C1 = "=IF(RC[-2]>(Ticket!R22C2),6,IF(RC[-2]>Ticket!R21C2,5,4))"
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    ActiveSheet.Range("F:F").AutoFilter Field:=6
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).FormulaR1C1 = "=IF(RC[-1]>(Ticket!R22C2),""REENQUADRAMENTO"",IF(RC[-1]>Ticket!R21C2,""RETENÇÃO"",""RELACIONAMENTO""))"
        Cells(lin_inicio, 11).FormulaR1C1 = "=IF(RC[-2]>(Ticket!R22C2),6,IF(RC[-2]>Ticket!R21C2,5,4))"
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="<>CONTATADO"
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="CONTATADO"
    ActiveSheet.Range("B:B").AutoFilter Field:=2, Criteria1:="EQUIPE ATACADO PLUS"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).FormulaR1C1 = "=IF(RC[-1]>(Ticket!R23C2),""REENQUADRAMENTO"",IF(RC[-1]>Ticket!R22C2,""RETENÇÃO"",""RELACIONAMENTO""))"
        Cells(lin_inicio, 11).FormulaR1C1 = "=IF(RC[-2]>(Ticket!R23C2),6,IF(RC[-2]>Ticket!R22C2,5,4))"
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    ActiveSheet.Range("F:F").AutoFilter Field:=6
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).FormulaR1C1 = "=IF(RC[-1]>(Ticket!R23C2),""REENQUADRAMENTO"",IF(RC[-1]>Ticket!R22C2,""RETENÇÃO"",""RELACIONAMENTO""))"
        Cells(lin_inicio, 11).FormulaR1C1 = "=IF(RC[-2]>(Ticket!R23C2),6,IF(RC[-2]>Ticket!R22C2,5,4))"
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="<>CONTATADO"
    ActiveSheet.Range("F:F").AutoFilter Field:=6, Criteria1:="CONTATADO"
    ActiveSheet.Range("B:B").AutoFilter Field:=2, Criteria1:="*EQUIPE MASTER*"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).FormulaR1C1 = "=IF(RC[-1]>(Ticket!R23C2*2),""REENQUADRAMENTO"",IF(RC[-1]>Ticket!R23C2,""RETENÇÃO"",""RELACIONAMENTO""))"
        Cells(lin_inicio, 11).FormulaR1C1 = "=IF(RC[-2]>(Ticket!R23C2*2),6,IF(RC[-2]>Ticket!R23C2,5,4))"
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    ActiveSheet.Range("F:F").AutoFilter Field:=6
    ActiveSheet.Range("D:D").AutoFilter Field:=4, Criteria1:="CONTATADO"
    lin_inicio = ActiveSheet.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    If Not IsEmpty(Cells(lin_inicio, 1).Value) Then
        Cells(lin_inicio, 10).FormulaR1C1 = "=IF(RC[-1]>(Ticket!R23C2*2),""REENQUADRAMENTO"",IF(RC[-1]>Ticket!R23C2,""RETENÇÃO"",""RELACIONAMENTO""))"
        Cells(lin_inicio, 11).FormulaR1C1 = "=IF(RC[-2]>(Ticket!R23C2*2),6,IF(RC[-2]>Ticket!R23C2,5,4))"
        If Not IsEmpty(Cells(lin_inicio, 1).End(xlDown).Value) Then
            Range("J" & lin_inicio & ":K" & Cells(lin_inicio, 1).End(xlDown).Row).SpecialCells(xlCellTypeVisible).FillDown
        End If
    End If
    On Error Resume Next
    ActiveSheet.ShowAllData
    
    Columns("J:K").Copy
    Columns("J:K").PasteSpecial Paste:=xlPasteValues
    
    Application.ScreenUpdating = True
    MsgBox "AJUSTAR QUANTIDADE DE COMPRA > 3 PARA CLIENTES COM STATUS 'PROSPECÇÃO'"
End Sub

