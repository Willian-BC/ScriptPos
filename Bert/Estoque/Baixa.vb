Private Sub CommandButton201_Click()
'OK BAIXA

    If OptionButton3.Value = True And TextBox201 = "KB" Then
    ElseIf OptionButton3.Value = True And TextBox201 <> "KB" Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "PEDIDO INFORMADO NÃO É KANBAN", vbCritical
        GoTo FIM
    ElseIf TextBox201 = "" Or TextBox202 = "" Or (OptionButton1.Value = False And OptionButton2.Value = False) Then
        MsgBox "FAVOR PREENCHER OS CAMPOS PEDIDO E CLIENTE!", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    Sheets("Base de dados").Visible = xlSheetVisible
    Sheets("Base de dados").Activate
    If Not Sheets("Base de dados").AutoFilterMode Then
        Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
    End If
    On Error Resume Next
    Sheets("Base de dados").ShowAllData
    
    Set rngAF = Range("A1:A" & Cells(1, 1).End(xlDown).Row)
    
    If TextBox201 <> "" Then
        Sheets("Base de dados").Range("D:D").AutoFilter Field:=4, Criteria1:=TextBox201.Text
    End If
    If TextBox202 <> "" Then
        Sheets("Base de dados").Range("G:G").AutoFilter Field:=7, Criteria1:="=*" & TextBox202.Text & "*"
    End If
    If TextBox203 <> "" Then
        Sheets("Base de dados").Range("C:C").AutoFilter Field:=3, Criteria1:="=*" & TextBox203.Text & "*"
    End If
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    If Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "PEDIDO INFORMADO NÃO EXISTE !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        Sheets("Base de dados").ShowAllData
        GoTo FIM
    End If
NOVO:
    Dim arrayItems2()
    With Planilha5
        ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, Sheets("Base de dados").Range("A:A")), 1 To 11) '.UsedRange.Columns.Count
        Me.ListBox2.ColumnCount = 11 '.UsedRange.Columns.Count
        Me.ListBox2.ColumnWidths = "40;130;150;80;90;80;90;90;80;80;200"
        i = 0
        For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
            Me.ListBox2.AddItem
            For coluna = 1 To 11 '.UsedRange.Columns.Count
                arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
            Next coluna
            i = i + 1
        Next rngcell
        Me.ListBox2.List = arrayItems2()
    End With

'''''''''''''''''''''''BAIXA COMPLETA'''''''''''''''''''''''''''''''
    If OptionButton1.Value = True Then
        result = MsgBox("TEM CERTEZA QUE DESEJA DAR BAIXA EM TODO O PEDIDO " & TextBox201 & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
        If result = vbYes Then
            Do While Cells(lin_inicio, 1).Value <> 0
                Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Baixa").Rows(Sheets("Baixa").Cells(1, 1).End(xlDown).Row + 1)
                Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 14) = Now
                Sheets("Base de dados").Rows(lin_inicio).Delete
                lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            Loop
            Sheets("Base de dados").ShowAllData
            Sheets("Base de dados").Cells(2, 1) = 1
            Sheets("Base de dados").Cells(3, 1) = 2
            Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 2).End(xlDown).Row)
        Else
            Sheets("Base de dados").ShowAllData
            GoTo FIM
        End If
        MsgBox "BAIXA REALIZADA COM SUCESSO !", vbInformation
        TextBox201 = ""
        TextBox202 = ""
        TextBox203 = ""
        OptionButton1.Value = False
        OptionButton2.Value = False
        OptionButton3.Value = False
        ListBox2.Clear
        GoTo FIM

'''''''''''''''''''''''BAIXA PARCIAL'''''''''''''''''''''''''''''''
    ElseIf OptionButton2.Value = True Then
AUX:
        ID = Application.InputBox("INFORME O ID")
        If ID = 0 Then
            Sheets("Base de dados").ShowAllData
            Sheets("Base de dados").Cells(2, 1) = 1
            Sheets("Base de dados").Cells(3, 1) = 2
            Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 2).End(xlDown).Row)
            GoTo FIM
        End If
        result = MsgBox("TEM CERTEZA QUE DESEJA DAR BAIXA NO ID " & ID & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
        If result = vbYes Then
            Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=ID
            lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            
            If Cells(lin_inicio, 1).Value = 0 Then
                MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR O NÚMERO DE ID", vbCritical
                Sheets("Base de dados").Range("A:A").AutoFilter Field:=1
                GoTo AUX
            End If
            
            Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Baixa").Rows(Sheets("Baixa").Cells(1, 1).End(xlDown).Row + 1)
            Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 14) = Now
            Sheets("Base de dados").Rows(lin_inicio).Delete
            
            result = MsgBox("BAIXA REALIZADA COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA BAIXAR OUTRO ITEM DESSE PEDIDO ?", vbYesNo + vbInformation)
            If result = vbYes Then
                Sheets("Base de dados").Range("A:A").AutoFilter Field:=1
                GoTo NOVO
            Else
                Sheets("Base de dados").ShowAllData
                Sheets("Base de dados").Cells(2, 1) = 1
                Sheets("Base de dados").Cells(3, 1) = 2
                Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 2).End(xlDown).Row)
                TextBox201 = ""
                TextBox202 = ""
                TextBox203 = ""
                OptionButton1.Value = False
                OptionButton2.Value = False
                OptionButton3.Value = False
                ListBox2.Clear
            End If
        Else
            Sheets("Base de dados").ShowAllData
            Sheets("Base de dados").Cells(2, 1) = 1
            Sheets("Base de dados").Cells(3, 1) = 2
            Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 2).End(xlDown).Row)
            GoTo FIM
        End If

'''''''''''''''''''''''BAIXA KANBAN'''''''''''''''''''''''''''''''
    ElseIf OptionButton3.Value = True Then
AUX2:
        ID = Application.InputBox("INFORME O ID")
        If ID = 0 Then
            Sheets("Base de dados").ShowAllData
            Sheets("Base de dados").Cells(2, 1) = 1
            Sheets("Base de dados").Cells(3, 1) = 2
            Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 2).End(xlDown).Row)
            GoTo FIM
        End If
        Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=ID
        lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        If Cells(lin_inicio, 1).Value = 0 Then
            MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR O NÚMERO DE ID", vbCritical
            Sheets("Base de dados").Range("A:A").AutoFilter Field:=1
            GoTo AUX2
        End If
ERRO:
        QTD = Application.InputBox("INFORME A QUANTIDADE QUE DESEJA DAR BAIXA NO ID " & ID & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA")
        If QTD > 0 Then
            If Sheets("Base de dados").Cells(lin_inicio, 8).Value - QTD > 0 Then
                Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Baixa").Rows(Sheets("Baixa").Cells(1, 1).End(xlDown).Row + 1)
                Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 8) = QTD
                Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 14) = Now
                Sheets("Base de dados").Cells(lin_inicio, 8).Value = Sheets("Base de dados").Cells(lin_inicio, 8).Value - QTD
            ElseIf Sheets("Base de dados").Cells(lin_inicio, 8).Value - QTD = 0 Then
                Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Baixa").Rows(Sheets("Baixa").Cells(1, 1).End(xlDown).Row + 1)
                Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 14) = Now
                Sheets("Base de dados").Rows(lin_inicio).Delete
            Else
                MsgBox "VALOR INFORMADO MAIOR QUE ESTOQUE" & vbCrLf & vbCrLf & "VERIFIQUE A QUANTIDADE INFORMADA"
                GoTo ERRO
            End If
            result = MsgBox("BAIXA REALIZADA COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA BAIXAR OUTRO ITEM ?", vbYesNo + vbInformation)
            If result = vbYes Then
                Sheets("Base de dados").Range("A:A").AutoFilter Field:=1
                GoTo NOVO
            Else
                Sheets("Base de dados").ShowAllData
                Sheets("Base de dados").Cells(2, 1) = 1
                Sheets("Base de dados").Cells(3, 1) = 2
                Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 1).End(xlDown).Row)
                TextBox201 = ""
                TextBox202 = ""
                TextBox203 = ""
                OptionButton1.Value = False
                OptionButton2.Value = False
                OptionButton3.Value = False
                ListBox2.Clear
            End If
        End If
    End If
FIM:
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
End Sub
