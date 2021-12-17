Sub sem_MP()

Dim confirma_peso As Integer
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.Calculation = xlManual
    
    StartTime = Timer
    
    ws = ThisWorkbook.ActiveSheet.Name
    
    'Limpa filtro da planilha ws
    On Error Resume Next
    Worksheets(ws).Range("A5").Select
    Worksheets(ws).ListObjects(1).ShowAutoFilter = True
    Worksheets(ws).ListObjects(1).AutoFilter.ShowAllData
    Worksheets(ws).Range("AA5").Select
    Worksheets(ws).ListObjects(2).ShowAutoFilter = True
    Worksheets(ws).ListObjects(2).AutoFilter.ShowAllData
    
    'Remove celulas com formulas
    If Cells(6, 18).HasFormula = True Or Cells(7, 19).HasFormula = True Then
        Worksheets(ws).Calculate
        Worksheets(ws).Range("R6:U" & Cells(5, 6).End(xlDown).Row).Copy
        Worksheets(ws).Range("R6:U" & Cells(5, 6).End(xlDown).Row).PasteSpecial Paste:=xlPasteValues
    End If
    
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
    
    'Retira itens sem data de componente da programação
    Worksheets(ws).Range("A5").Select
    Worksheets(ws).ListObjects(1).Range.AutoFilter Field:=22, Criteria1:=""
    Worksheets(ws).ListObjects(1).Range.AutoFilter Field:=23, Criteria1:="<>MP NOK"
    lin_inicio = Worksheets(ws).AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = Worksheets(ws).Cells(5, 6).End(xlDown).Row
    If Cells(lin_inicio, 6).Value <> 0 Then
        Worksheets(ws).Cells(lin_inicio, 23).Value = "MP NOK"
        Worksheets(ws).Range("W" & lin_inicio & ":W" & lin_fim).FillDown
    End If
    Worksheets(ws).ListObjects(1).AutoFilter.ShowAllData
    w = 6
    aux = 1
    
    'Limpa filtro aba programação
    Worksheets("Programação").Activate
    Worksheets("Programação").Range("A1").Select
    Worksheets("Programação").ListObjects(1).ShowAutoFilter = True
    Worksheets("Programação").ListObjects(1).AutoFilter.ShowAllData
    
    Worksheets(ws).Activate
    
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
                        If Worksheets("Programação").Cells(2, 1) = "" Then
                            y = 2
                        Else
                            y = Worksheets("Programação").Cells(1, 1).End(xlDown).Row + 1
                        End If
                        If aux1 = "" Then aux1 = y
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
                        If Worksheets("Programação").Cells(2, 1) = "" Then
                            y = 2
                        Else
                            y = Worksheets("Programação").Cells(1, 1).End(xlDown).Row + 1
                        End If
                        If aux1 = "" Then aux1 = y
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
                    If Worksheets(ws).Cells(k, 22) <> 0 And Worksheets(ws).Cells(k, 22) <> "" Then
                        Worksheets(ws).ListObjects(2).Range.AutoFilter Field:=6, Criteria1:=">" & Format(Worksheets(ws).Cells(k, 22), "mm/dd/yyyy")
                    End If
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
    If y = "" Then GoTo kill
    
    Worksheets("Programação").Activate
    Worksheets("Programação").Range("A1").Select
    Worksheets("Programação").ListObjects(1).ShowAutoFilter = True
    Worksheets("Programação").ListObjects(1).AutoFilter.ShowAllData
    
    Worksheets("Programação").Cells(aux1, 17).Formula2R1C1 = "=ISOWEEKNUM(RC[-15])"
    Worksheets("Programação").Cells(aux1, 18).Formula2R1C1 = "=PROPER(TEXT(RC[-16],""mmmm""))"
    Worksheets("Programação").Cells(aux1, 19).Formula2R1C1 = "=VLOOKUP(RC[-12],COMPONENTES!C[-18]:C[11],30,0)"
    Worksheets("Programação").Cells(aux1, 20).Formula2R1C1 = "=VLOOKUP(RC[-13],COMPONENTES!C[-19]:C[4],24,0)"
    Worksheets("Programação").Cells(aux1, 21).Formula2R1C1 = "=IF(AND(RC[-2]=""Produto Acabado"",RIGHT(RC[-13],3)=""398""),""Galvanizado"",IF(AND(RC[-2]=""Produto Acabado"",RIGHT(RC[-13],3)<>""398""),""Pintado"",""""))"
    Worksheets("Programação").Cells(aux1, 22).Formula2R1C1 = "=VLOOKUP(RC[-15],COMPONENTES!C[-21]:C[-20],2,0)"
    Worksheets("Programação").Cells(aux1, 23).Formula2R1C1 = "=IF(OR(RC[-18]=0,RC[-18]=302),""Kanban"",XLOOKUP(RC[-18],'COMPONENTES (4PR)'!C[-16],'COMPONENTES (4PR)'!C[-8],"""",0,1))"
    Worksheets("Programação").Cells(aux1, 24).Formula2R1C1 = "=RC[-16]&RC[-19]"
    
    Worksheets("Programação").Range("Q" & aux1 & ":X" & y).FillDown
    Worksheets("Programação").Calculate
    Worksheets("Programação").Range("Q" & aux1 & ":X" & y).Copy
    Worksheets("Programação").Range("Q" & aux1 & ":X" & y).PasteSpecial Paste:=xlPasteValues
kill:
    Worksheets(ws).Activate
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    MsgBox "Concluído em: " & Format((Timer - StartTime) / 86400, "hh:mm:ss")
    
End Sub

-------------------------------------------------------

Sub programar_manual()

Dim confirma_peso As Integer
    
    Application.ScreenUpdating = False
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    ActiveSheet.DisplayPageBreaks = False
    Application.Calculation = xlManual
    
    StartTime = Timer
    
    ws = ThisWorkbook.ActiveSheet.Name
    
    'Limpa filtro da planilha ws
    On Error Resume Next
    Worksheets(ws).Range("A5").Select
    Worksheets(ws).ListObjects(1).ShowAutoFilter = True
    Worksheets(ws).ListObjects(1).AutoFilter.ShowAllData
    Worksheets(ws).Range("AA5").Select
    Worksheets(ws).ListObjects(2).ShowAutoFilter = True
    Worksheets(ws).ListObjects(2).AutoFilter.ShowAllData
    
    'Remove celulas com formulas
    If Cells(6, 18).HasFormula = True Or Cells(7, 19).HasFormula = True Then
        Worksheets(ws).Calculate
        Worksheets(ws).Range("R6:U" & Cells(5, 6).End(xlDown).Row).Copy
        Worksheets(ws).Range("R6:U" & Cells(5, 6).End(xlDown).Row).PasteSpecial Paste:=xlPasteValues
    End If
    
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
    
    'Retira itens sem data de componente da programação
    Worksheets(ws).Range("A5").Select
    Worksheets(ws).ListObjects(1).Range.AutoFilter Field:=22, Criteria1:=""
    Worksheets(ws).ListObjects(1).Range.AutoFilter Field:=23, Criteria1:="<>MP NOK"
    lin_inicio = Worksheets(ws).AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = Worksheets(ws).Cells(5, 6).End(xlDown).Row
    If Cells(lin_inicio, 6).Value <> 0 Then
        Worksheets(ws).Cells(lin_inicio, 23).Value = "MP NOK"
        Worksheets(ws).Range("W" & lin_inicio & ":W" & lin_fim).FillDown
    End If
    Worksheets(ws).ListObjects(1).AutoFilter.ShowAllData
    w = 6
    aux = 1
    
    'Limpa filtro aba programação
    Worksheets("Programação").Activate
    Worksheets("Programação").Range("A1").Select
    Worksheets("Programação").ListObjects(1).ShowAutoFilter = True
    Worksheets("Programação").ListObjects(1).AutoFilter.ShowAllData
    
    Worksheets(ws).Activate
    
    'Classificar manual kamban
    With Worksheets(ws).ListObjects(1).Sort
       .SortFields.Clear
       .SortFields.Add Key:=Range("W5"), SortOn:=xlSortOnValues, Order:=xlAscending
       .Header = xlYes
       .Apply
    End With
    
    For k = 6 To num_linhas
inicio:
        If Worksheets(ws).Cells(k, 11).Value > 0 And IsNumeric(Worksheets(ws).Cells(k, 23).Value) Then
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
            Else
                Worksheets(ws).Cells(k, 23).Value = "MP NOK"
                GoTo prox
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
                        If Worksheets("Programação").Cells(2, 1) = "" Then
                            y = 2
                        Else
                            y = Worksheets("Programação").Cells(1, 1).End(xlDown).Row + 1
                        End If
                        If aux1 = "" Then aux1 = y
                        'Copia ordem de produção
                        Worksheets("Programação").Cells(y, 1) = ws
                        Worksheets("Programação").Cells(y, 2).Value = Worksheets(ws).Cells(w, 32).Value
                        Worksheets(ws).Range("A" & k & ":J" & k).Copy Worksheets("Programação").Cells(y, 3)
                        Worksheets("Programação").Cells(y, 13) = peso

                        Worksheets("Programação").Cells(y, 14) = peso / peso_unitário
                        Worksheets("Programação").Cells(y, 15) = (Worksheets("Programação").Cells(y, 13) / Worksheets("Programação").Cells(y, 11))
                        Worksheets("Programação").Cells(y, 16) = "Processado"
                        
                        Worksheets(ws).Activate
                        
                    ElseIf (Worksheets(ws).Cells(w, 30).Value - Worksheets(ws).Cells(k, 11).Value) < 0 Then
                        peso = Worksheets(ws).Cells(w, 30).Value
                        Worksheets(ws).Cells(k, 11).Value = Worksheets(ws).Cells(k, 11).Value - Worksheets(ws).Cells(w, 30).Value
                        Worksheets(ws).Cells(w, 33).Value = Worksheets(ws).Cells(w, 33).Value + Worksheets(ws).Cells(w, 30).Value
                        Worksheets(ws).Cells(w, 30).Value = 0
                        peso_unitário = (Worksheets(ws).Cells(k, 9).Value / Worksheets(ws).Cells(k, 10).Value)
                        
                        Worksheets("Programação").Activate
                        If Worksheets("Programação").Cells(2, 1) = "" Then
                            y = 2
                        Else
                            y = Worksheets("Programação").Cells(1, 1).End(xlDown).Row + 1
                        End If
                        If aux1 = "" Then aux1 = y
                        'Copia ordem de produção
                        Worksheets("Programação").Cells(y, 1) = ws
                        Worksheets("Programação").Cells(y, 2).Value = Worksheets(ws).Cells(w, 32).Value
                        Worksheets(ws).Range("A" & k & ":J" & k).Copy Worksheets("Programação").Cells(y, 3)
                        Worksheets("Programação").Cells(y, 13) = peso
                        Worksheets("Programação").Cells(y, 14) = peso / peso_unitário
                        Worksheets("Programação").Cells(y, 15) = (Worksheets("Programação").Cells(y, 13) / Worksheets("Programação").Cells(y, 11))
                        Worksheets("Programação").Cells(y, 16) = "Processado"
                        
                        w = w + 1
                        Worksheets(ws).Activate
                        GoTo continua
                    End If
                Else
                    Worksheets(ws).Range("AA5").Select
                    Worksheets(ws).ListObjects(2).Range.AutoFilter Field:=4, Criteria1:=">0"
                    If Worksheets(ws).Cells(k, 22) <> 0 And Worksheets(ws).Cells(k, 22) <> "" Then
                        Worksheets(ws).ListObjects(2).Range.AutoFilter Field:=6, Criteria1:=">" & Format(Worksheets(ws).Cells(k, 22), "mm/dd/yyyy")
                    End If
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
            Worksheets(ws).ListObjects(1).Range.AutoFilter Field:=23, Criteria1:=">=-1000", Operator:=xlAnd, Criteria2:="<=9999"
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
    
    'Reclassificar manual kamban
    With Worksheets(ws).ListObjects(1).Sort
       .SortFields.Clear
       .SortFields.Add Key:=Range("A5"), SortOn:=xlSortOnValues, Order:=xlAscending
       .SortFields.Add Key:=Range("B5"), SortOn:=xlSortOnValues, Order:=xlAscending
       .SortFields.Add Key:=Range("E5"), SortOn:=xlSortOnValues, Order:=xlAscending
       .Header = xlYes
       .Apply
    End With
    
    If y = "" Then GoTo kill
    
    Worksheets("Programação").Activate
    Worksheets("Programação").Range("A1").Select
    Worksheets("Programação").ListObjects(1).ShowAutoFilter = True
    Worksheets("Programação").ListObjects(1).AutoFilter.ShowAllData
    
    Worksheets("Programação").Cells(aux1, 17).Formula2R1C1 = "=ISOWEEKNUM(RC[-15])"
    Worksheets("Programação").Cells(aux1, 18).Formula2R1C1 = "=PROPER(TEXT(RC[-16],""mmmm""))"
    Worksheets("Programação").Cells(aux1, 19).Formula2R1C1 = "=VLOOKUP(RC[-12],COMPONENTES!C[-18]:C[11],30,0)"
    Worksheets("Programação").Cells(aux1, 20).Formula2R1C1 = "=VLOOKUP(RC[-13],COMPONENTES!C[-19]:C[4],24,0)"
    Worksheets("Programação").Cells(aux1, 21).Formula2R1C1 = "=IF(AND(RC[-2]=""Produto Acabado"",RIGHT(RC[-13],3)=""398""),""Galvanizado"",IF(AND(RC[-2]=""Produto Acabado"",RIGHT(RC[-13],3)<>""398""),""Pintado"",""""))"
    Worksheets("Programação").Cells(aux1, 22).Formula2R1C1 = "=VLOOKUP(RC[-15],COMPONENTES!C[-21]:C[-20],2,0)"
    Worksheets("Programação").Cells(aux1, 23).Formula2R1C1 = "=IF(OR(RC[-18]=0,RC[-18]=302),""Kanban"",XLOOKUP(RC[-18],'COMPONENTES (4PR)'!C[-16],'COMPONENTES (4PR)'!C[-8],"""",0,1))"
    Worksheets("Programação").Cells(aux1, 24).Formula2R1C1 = "=RC[-16]&RC[-19]"
    
    Worksheets("Programação").Range("Q" & aux1 & ":X" & y).FillDown
    Worksheets("Programação").Calculate
    Worksheets("Programação").Range("Q" & aux1 & ":X" & y).Copy
    Worksheets("Programação").Range("Q" & aux1 & ":X" & y).PasteSpecial Paste:=xlPasteValues
kill:
    Worksheets(ws).Activate
    Application.ScreenUpdating = True
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    ActiveSheet.DisplayPageBreaks = True
    MsgBox "Concluído em: " & Format((Timer - StartTime) / 86400, "hh:mm:ss")
    
End Sub

