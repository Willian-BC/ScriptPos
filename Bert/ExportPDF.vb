Sub PDF()
  
'''''''''''''''''''''''''''EXPORTAR DE TABELA DINAMICA'''''''''''''''''''''''''''
    Application.ScreenUpdating = False
    Dim saveLocation As String
    Dim rng As Range
    Dim ws As Worksheet
    Dim xPTable As PivotTable
    Dim xPFile As PivotField
    On Error Resume Next
    
'''''''''''''''''''''''''''CRIA PASTA PARA RECEBER OS ARQUIVOS'''''''''''''''''''''''''''
    saveLocation = "C:\PDF\"
    If Dir(saveLocation, vbDirectory) = "" Then MkDir saveLocation
         
    For i = 1 To 15
        If i = 1 Then
            maq = "Pintar"
            k = 1
            Set ws = Sheets("Perfilamento")
            Set xPTable = ws.PivotTables("Tabela dinâmica1")
            Set xPFile1 = xPTable.PivotFields("Comp. (S/N)")
            Set xPFile2 = xPTable.PivotFields("Tipo da OP")
            Set xPFile3 = xPTable.PivotFields("MP (OK/NOK)")
            Set xPFile4 = xPTable.PivotFields("Operação Comp.")
            Set xPFile5 = xPTable.PivotFields("Acab.")
            
'''''''''''''''''''''''''''LIMPA TODOS OS FILTROS'''''''''''''''''''''''''''
            xPTable.ClearAllFilters

'''''''''''''''''''''''''''FILTRO EM CADA CAMPO'''''''''''''''''''''''''''
            With xPFile1
                .PivotItems("Componente").Visible = False
                .PivotItems("Pedido").Visible = True
            End With
            With xPFile2
                .PivotItems("Kanban").Visible = True
                .PivotItems("Minuta").Visible = True
                .PivotItems("Pedido").Visible = True
                .PivotItems("Plano Blank").Visible = False
                .PivotItems("Plano Sliter").Visible = False
                .PivotItems("Verificar").Visible = False
            End With

'''''''''''''''''''''''''''FILTRA VALOR UNICO'''''''''''''''''''''''''''
            xPFile3.PivotFilters.Add2 xlCaptionEquals, Value1:="OK"
        ElseIf i = 2 Then
            maq = "Daltec"
            xPFile5.ClearAllFilters
        ElseIf i = 3 Then maq = "Fiorentini"
        ElseIf i = 4 Then maq = "Progressivas"
        ElseIf i = 5 Then maq = "Corte Guilhotina"
        ElseIf i = 6 Then maq = "Corte Perfis e Tubos"
        ElseIf i = 7 Then maq = "Estampar"
        ElseIf i = 8 Then maq = "Dobrar"
        ElseIf i = 9 Then maq = "Estampar/Dobrar"
        ElseIf i = 10 Then maq = "Soldar"
        ElseIf i = 11 Then maq = "Embalagem MDP BMA"
        ElseIf i = 12 Then maq = "Terceiro"
        ElseIf i = 13 Then maq = "Embalar PA"
        ElseIf i = 14 Then maq = "Corte MDP BMA"
        ElseIf i = 15 Then maq = "Zikelis"
        End If
                
'''''''''''''''''''''''''''LIMPA FILTRO DO CAMPO ESPECIFICO'''''''''''''''''''''''''''
        xPFile4.ClearAllFilters
        For j = 1 To xPFile4.PivotItems.Count
            If xPFile4.PivotItems(j).Name = maq Then
                xPFile4.PivotItems(xPFile4.PivotItems(j).Name).Visible = True
            Else
                xPFile4.PivotItems(xPFile4.PivotItems(j).Name).Visible = False
            End If
        Next
        ws.Cells(6, 6) = UCase(maq)
        If i = 9 Then maq = "Estampar-Dobrar"
        
novo:
        If i = 1 And k = 1 Then
            xPFile5.ClearAllFilters
            acab = "AM"
            xPFile5.PivotFilters.Add2 xlCaptionContains, Value1:=acab
        ElseIf i = 1 And k = 2 Then
            xPFile5.ClearAllFilters
            acab = "AZ"
            xPFile5.PivotFilters.Add2 xlCaptionContains, Value1:=acab
        ElseIf i = 1 And k = 3 Then
            xPFile5.ClearAllFilters
            acab = "CZ"
            xPFile5.PivotFilters.Add2 xlCaptionContains, Value1:=acab
        ElseIf i = 1 And k = 4 Then
            xPFile5.ClearAllFilters
            acab = "LA"
            xPFile5.PivotFilters.Add2 xlCaptionContains, Value1:=acab
        ElseIf i = 1 And k = 5 Then
            xPFile5.ClearAllFilters
            For j = 1 To xPFile5.PivotItems.Count
                If xPFile5.PivotItems(j).Name = " AM" Or xPFile5.PivotItems(j).Name = " AZ" Or _
                xPFile5.PivotItems(j).Name = " CZ" Or xPFile5.PivotItems(j).Name = " LA" Then
                    xPFile5.PivotItems(xPFile5.PivotItems(j).Name).Visible = False
                Else
                    xPFile5.PivotItems(xPFile5.PivotItems(j).Name).Visible = True
                End If
            Next
        End If
        
        fim = ws.Cells(14, 10).Offset(1)
        If fim = "" Then GoTo prox
        fim = ws.Cells(14, 10).End(xlDown).Row + 1
        Set rng = ws.Range("A3:N" & fim)
        
        If i = 1 And k = 1 Then
                
'''''''''''''''''''''''''''DEFINIÇÃO DAS CONFIGURAÇÕES DE IMPRESSÃO'''''''''''''''''''''''''''
        Application.PrintCommunication = False
        With ws.PageSetup
            .Zoom = False
            .BlackAndWhite = False
            .FitToPagesTall = False
            .FitToPagesWide = 1
            .CenterHorizontally = True
            .PaperSize = xlPaperA4
            .Orientation = xlPortrait
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.25)
            .FooterMargin = Application.InchesToPoints(0.25)
        End With
        Application.PrintCommunication = True
        End If
prox:
        If i = 1 And k < 5 Then

'''''''''''''''''''''''''''EXCLUIR ARQUIVO EXISTENTE COM MESMO NOME'''''''''''''''''''''''''''
            Kill saveLocation & maq & " - " & acab & ".pdf"
                
'''''''''''''''''''''''''''EXPORTAR ARQUIVO EM PDF'''''''''''''''''''''''''''
            If fim <> "" Then rng.ExportAsFixedFormat Type:=xlTypePDF, Filename:=saveLocation & maq & " - " & acab & ".pdf"
            k = k + 1
            GoTo novo
        Else
            Kill saveLocation & maq & ".pdf"
            If fim <> "" Then rng.ExportAsFixedFormat Type:=xlTypePDF, Filename:=saveLocation & maq & ".pdf"
        End If
    Next
    xPFile3.ClearAllFilters
    Application.ScreenUpdating = True
    MsgBox "Finalizado com sucesso", vbExclamation
End Sub
