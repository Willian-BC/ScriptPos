Sub sem_MP()

Dim confirma_peso As Integer
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.Calculation = xlManual
    
    tempo = Now()
    
    ws = ThisWorkbook.ActiveSheet.Name
    
    'Limpa filtro da planilha ws
    On Error Resume Next
    Worksheets(ws).Range("A5").Select
    Worksheets(ws).ListObjects(1).ShowAutoFilter = True
    Worksheets(ws).ListObjects(1).AutoFilter.ShowAllData
    Worksheets(ws).Range("AA5").Select
    Worksheets(ws).ListObjects(2).ShowAutoFilter = True
    Worksheets(ws).ListObjects(2).AutoFilter.ShowAllData
    
    ''''Solicita confirmação de limpeza do peso restante
    confirma_peso = MsgBox("Limpar coluna peso restante?", vbYesNoCancel + vbQuestion + vbDefaultButton3)
    
    If confirma_peso = vbCancel Then
        Exit Sub
    ElseIf confirma_peso = vbYes Then
        Worksheets(ws).Columns(11).ClearContents
        Worksheets(ws).Columns("I:I").Copy Worksheets(ws).Columns("K:K")
        Worksheets(ws).Cells(5, 11) = "Peso Restante"
    End If
    
    'conta número de ordens de produção
    num_linhas = Worksheets(ws).Cells(5, 6).End(xlDown).Row
    
    'verifica falta de programação
    Worksheets(ws).Range("AD5").Select
    Worksheets(ws).ListObjects(2).Range.AutoFilter Field:=4, Criteria1:=">0"
    lin_inicio = Worksheets(ws).AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    Worksheets(ws).ListObjects(2).AutoFilter.ShowAllData
    
    If Cells(lin_inicio, 30).Value = 0 Then
        verifica_capacidade = MsgBox("Falta Capacidade Produtiva", vbCritical) = vbOK
        GoTo fim
    End If
    
    'conta número de dias disponpíveis para programar
    num_linhas_prog = Worksheets(ws).Cells(6, 32).End(xlDown).Row
    
    w = 6
    aux = 1
    
    For k = 6 To num_linhas
inicio:
        If Worksheets(ws).Cells(k, 11).Value > 0 And Worksheets(ws).Cells(k, 23).Value = "" Then
            If Worksheets(ws).Cells(k, 22) <> 0 And Worksheets(ws).Cells(k, 22) <> "" Then
                Worksheets(ws).Range("AA5").Select
                Worksheets(ws).ListObjects(2).Range.AutoFilter Field:=4, Criteria1:=">0"
                Worksheets(ws).ListObjects(2).Range.AutoFilter Field:=6, Criteria1:=">" & Format(Worksheets(ws).Cells(k, 22), "mm/dd/yyyy")
                lin_inicio = Worksheets(ws).AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
                Worksheets(ws).ListObjects(2).AutoFilter.ShowAllData
                If Cells(lin_inicio, 32).Value = 0 Then
                    Worksheets(ws).Cells(k, 23).Value = "NP"
                    GoTo prox
                End If
                If w < lin_inicio Then
                    aux = 10
                End If
                w = lin_inicio
            ElseIf aux = 10 Then
                Worksheets(ws).Range("AA5").Select
                Worksheets(ws).ListObjects(2).Range.AutoFilter Field:=4, Criteria1:=">0"
                lin_inicio = Worksheets(ws).AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
                Worksheets(ws).ListObjects(2).AutoFilter.ShowAllData
                w = lin_inicio
                aux = 1
            End If
continua:
            peso = Worksheets(ws).Cells(k, 11).Value
            If w <= num_linhas_prog Then
                If Worksheets(ws).Cells(w, 30).Value > 0 Then
                    If (Worksheets(ws).Cells(w, 30).Value - Worksheets(ws).Cells(k, 11).Value) >= 0 Then
                        Worksheets(ws).Cells(w, 33).Value = Worksheets(ws).Cells(w, 33).Value + Worksheets(ws).Cells(k, 11).Value
                        Worksheets(ws).Cells(w, 30).Value = Worksheets(ws).Cells(w, 30).Value - Worksheets(ws).Cells(k, 11).Value
                        Worksheets(ws).Cells(k, 11).Value = 0
                        peso_unitário = (Worksheets(ws).Cells(k, 9).Value / Worksheets(ws).Cells(k, 10).Value)
                        
                        Worksheets("Programação").Activate
                        If Cells(2, 1) = "" Then
                            y = 2
                        Else
                            y = Cells(1, 1).End(xlDown).Row + 1
                        End If
                        
                        'Copia ordem de produção
                        Worksheets("Programação").Cells(y, 1) = ws
                        Worksheets("Programação").Cells(y, 2).Value = Worksheets(ws).Cells(w, 32).Value
                        Worksheets(ws).Range("A" & k & ":J" & k).Copy Worksheets("Programação").Cells(y, 3)
                        Worksheets("Programação").Cells(y, 13) = peso
                        Worksheets("Programação").Cells(y, 14) = peso / peso_unitário
                        Worksheets("Programação").Cells(y, 15) = (Worksheets("Programação").Cells(y, 13) / Worksheets("Programação").Cells(y, 11))
                        Worksheets("Programação").Cells(y, 16) = "Sem MP"
                        
                        Worksheets(ws).Activate
                        
                    ElseIf (Worksheets(ws).Cells(w, 30).Value - Worksheets(ws).Cells(k, 11).Value) < 0 Then
                        peso = Worksheets(ws).Cells(w, 30).Value
                        Worksheets(ws).Cells(k, 11).Value = Worksheets(ws).Cells(k, 11).Value - Worksheets(ws).Cells(w, 30).Value
                        Worksheets(ws).Cells(w, 33).Value = Worksheets(ws).Cells(w, 33).Value + Worksheets(ws).Cells(w, 30).Value
                        Worksheets(ws).Cells(w, 30).Value = 0
                        peso_unitário = (Worksheets(ws).Cells(k, 9).Value / Worksheets(ws).Cells(k, 10).Value)
                        
                        Worksheets("Programação").Activate
                        If Cells(2, 1) = "" Then
                            y = 2
                        Else
                            y = Cells(1, 1).End(xlDown).Row + 1
                        End If
                        
                        'Copia ordem de produção
                        Worksheets("Programação").Cells(y, 1) = ws
                        Worksheets("Programação").Cells(y, 2).Value = Worksheets(ws).Cells(w, 32).Value
                        Worksheets(ws).Range("A" & k & ":J" & k).Copy Worksheets("Programação").Cells(y, 3)
                        Worksheets("Programação").Cells(y, 13) = peso
                        Worksheets("Programação").Cells(y, 14) = peso / peso_unitário
                        Worksheets("Programação").Cells(y, 15) = (Worksheets("Programação").Cells(y, 13) / Worksheets("Programação").Cells(y, 11))
                        Worksheets("Programação").Cells(y, 16) = "Sem MP"
                        
                        w = w + 1
                        Worksheets(ws).Activate
                        GoTo continua
                    End If
                Else
                    Worksheets(ws).Range("AA5").Select
                    Worksheets(ws).ListObjects(2).Range.AutoFilter Field:=4, Criteria1:=">0"
                    lin_inicio = Worksheets(ws).AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
                    Worksheets(ws).ListObjects(2).AutoFilter.ShowAllData
                    If Cells(lin_inicio, 30).Value = 0 Then
                        MsgBox "Fim da Capacidade Produtiva", vbCritical = vbOK
                        GoTo fim
                    End If
                    w = lin_inicio
                    GoTo continua
                End If
            Else
                MsgBox "Fim da Capacidade Produtiva", vbCritical = vbOK
                GoTo fim
            End If
        Else
            Worksheets(ws).Range("A5").Select
            Worksheets(ws).ListObjects(1).Range.AutoFilter Field:=11, Criteria1:=">0"
            Worksheets(ws).ListObjects(1).Range.AutoFilter Field:=23, Criteria1:=""
            lin_inicio = Worksheets(ws).AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            Worksheets(ws).ListObjects(1).AutoFilter.ShowAllData
            If Cells(lin_inicio, 11).Value = 0 Then
                MsgBox "Nenhum item para programar", vbCritical = vbOK
                GoTo fim
            End If
            k = lin_inicio
            GoTo inicio
        End If
prox:
    Next
fim:
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    execucao = Format((Now() - tempo) * 24 * 3600, "#0")
    respostafinal = MsgBox("Concluído em: " & execucao & "s", vbInformation) = vbOK
    
End Sub
