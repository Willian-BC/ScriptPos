Private Sub UserForm_Initialize()
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    Sheets("Baixa").Visible = xlSheetVeryHidden
    Sheets("Excluido").Visible = xlSheetVeryHidden
    Sheets("Descrição").Visible = xlSheetVeryHidden
End Sub

Private Sub CommandButton100_Click()
    Dim MyValue As Variant
    MyValue = InputBox("Digite a senha")
    If MyValue = "1010" Then
        Sheets("Base de dados").Visible = xlSheetVisible
        Sheets("Baixa").Visible = xlSheetVisible
        Sheets("Excluido").Visible = xlSheetVisible
        Sheets("Descrição").Visible = xlSheetVisible
    Else
        MsgBox ("Senha Incorreta")
    End If
End Sub

Private Sub CommandButton200_Click()
    ThisWorkbook.Save
End Sub

Private Sub CommandButton1_Click()
'OK PESQUISA'
    Application.ScreenUpdating = False
    Sheets("Base de dados").Visible = xlSheetVisible
    Sheets("Base de dados").Activate
    Set wb = ActiveWorkbook
    
    If Not Sheets("Base de dados").AutoFilterMode Then
        Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
    End If
    On Error Resume Next
    Sheets("Base de dados").ShowAllData
    
    Set rngAF = Range("A1:A" & Cells(1, 1).End(xlDown).Row)
    
    If TextBox1 <> "" Then Sheets("Base de dados").Range("D:D").AutoFilter Field:=4, Criteria1:="=*" & TextBox1.Text & "*"
    If TextBox2 <> "" Then Sheets("Base de dados").Range("E:E").AutoFilter Field:=5, Criteria1:=TextBox2.Text
    If TextBox3 <> "" Then Sheets("Base de dados").Range("F:F").AutoFilter Field:=6, Criteria1:="=*" & TextBox3.Text & "*"
    If TextBox4 <> "" Then Sheets("Base de dados").Range("B:B").AutoFilter Field:=2, Criteria1:=TextBox4.Text
    If TextBox5 <> "" Then Sheets("Base de dados").Range("C:C").AutoFilter Field:=3, Criteria1:=TextBox5.Text
    If TextBox6 <> "" Then Sheets("Base de dados").Range("I:I").AutoFilter Field:=9, Criteria1:=TextBox6.Text
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = Sheets("Base de dados").Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        Sheets("Base de dados").ShowAllData
        GoTo FIM
    Else
    
    Dim arrayItems2()
        With Planilha5
            ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, Sheets("Base de dados").Range("A:A")), 1 To 10) '.UsedRange.Columns.Count
            Me.ListBox1.ColumnCount = 10 '.UsedRange.Columns.Count
            Me.ListBox1.ColumnWidths = "40;90;80;100;100;180;40;50;80;200"
            i = 0
            For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
                Me.ListBox1.AddItem
                For coluna = 1 To 10 '.UsedRange.Columns.Count
                    arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
                Next coluna
                i = i + 1
            Next rngcell
            Me.ListBox1.List = arrayItems2()
        End With
    End If
    
    If CheckBox1 = True Then
        Set rngAJ = Range("B1:K" & lin_fim).SpecialCells(xlCellTypeVisible)
        rngAJ.Copy
        Workbooks.Add
        Range("A1").PasteSpecial Paste:=xlPasteValues
        Columns("J:J").NumberFormat = "dd/mm/yyyy"
    End If
    
FIM:
    wb.Sheets("Base de dados").ShowAllData
    wb.Sheets("Base de dados").Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
    
    If CheckBox1 = True Then
        result = MsgBox("DADOS EXPORTADOS COM SUCESSO !" & vbCrLf & "DESEJA FECHAR O FORMULÁRIO ?" & vbCrLf & vbCrLf & "É NECESSÁRIO FECHAR PARA EDITAR OS DADOS", vbYesNo + vbInformation)
        If result = vbYes Then
            Unload Me
        Else
            ActiveWindow.WindowState = xlMinimized
            wb.Activate
            CheckBox1 = False
        End If
    End If
    
End Sub

Private Sub CommandButton2_Click()
'LIMPAR PESQUISA
    TextBox1 = ""
    TextBox2 = ""
    TextBox3 = ""
    TextBox4 = ""
    TextBox5 = ""
    TextBox6 = ""
    ListBox1.Clear
    CheckBox1 = False
End Sub

Private Sub CommandButton3_Click()
'CANCELAR PESQUISA
    result = MsgBox("DESEJA SALVAR AS ALTERAÇÕES ?", vbYesNo + vbCritical)
    If result = vbYes Then
        ThisWorkbook.Save
    End If
    End
End Sub

Private Sub CommandButton4_Click()
'EDITAR PESQUISA
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Sheets("Base de dados").Visible = xlSheetVisible
    Sheets("Base de dados").Activate
    
    ID = Application.InputBox("INFORME O ID")
    If ID = 0 Then
        GoTo FIM
    End If
    
    If Not Sheets("Base de dados").AutoFilterMode Then
        Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
    End If
    On Error Resume Next
    Sheets("Base de dados").ShowAllData
    
    Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=ID
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    TextBox101 = Sheets("Base de dados").Cells(lin_inicio, 1)
    TextBox102 = Sheets("Base de dados").Cells(lin_inicio, 2)
    TextBox103 = Sheets("Base de dados").Cells(lin_inicio, 3)
    TextBox104 = Sheets("Base de dados").Cells(lin_inicio, 4)
    TextBox105 = Sheets("Base de dados").Cells(lin_inicio, 5)
    TextBox106 = Sheets("Base de dados").Cells(lin_inicio, 6)
    larray = Split(Sheets("Base de dados").Cells(lin_inicio, 7), "/")
    TextBox1071 = larray(0)
    TextBox1072 = larray(1)
    TextBox108 = Sheets("Base de dados").Cells(lin_inicio, 8)
    TextBox109 = Sheets("Base de dados").Cells(lin_inicio, 9)
    TextBox110 = Sheets("Base de dados").Cells(lin_inicio, 10)
    
    Sheets("Base de dados").ShowAllData
    
    UserForm1.MultiPage1.Value = 1
FIM:
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton5_Click()
'APAGAR PESQUISA
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Sheets("Base de dados").Visible = xlSheetVisible
    Sheets("Base de dados").Activate
    
    result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR UM REGISTRO?", vbYesNo + vbCritical)
    If result = vbYes Then
        ID = Application.InputBox("INFORME O ID")
        If ID = 0 Then
            GoTo FIM
        End If
        result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR O ID " & ID & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
        If result = vbYes Then
            If Not Sheets("Base de dados").AutoFilterMode Then
                Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
            End If
            On Error Resume Next
            Sheets("Base de dados").ShowAllData
            
            Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=ID
            
            lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            
            Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Excluido").Rows(Sheets("Excluido").Cells(1, 1).End(xlDown).Row + 1)
            Sheets("Excluido").Cells(Sheets("Excluido").Cells(1, 1).End(xlDown).Row, 13) = Now
            Sheets("Base de dados").Rows(lin_inicio).Delete
            
            Sheets("Base de dados").ShowAllData
            
            Sheets("Base de dados").Cells(2, 1) = 1
            Sheets("Base de dados").Cells(3, 1) = 2
            Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 1).End(xlDown).Row)
            
            MsgBox "CADASTRO EXCLUIDO COM SUCESSO!", vbInformation
            ListBox1.Clear
        End If
    End If
FIM:
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton101_Click()
'OK CADASTRO'
    If TextBox101 <> "" Then
        MsgBox "CADASTRO JÁ EXISTE!" & vbCrLf & vbCrLf & "CLIQUE EM ATUALIZAR REGISTRO OU LIMPAR", vbCritical
        GoTo FIM
    ElseIf TextBox104 = "" Or TextBox1071 = "" Or TextBox1072 = "" Or TextBox108 = "" Or TextBox109 = "" Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS PARA CADASTRO!", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    Sheets("Base de dados").Visible = xlSheetVisible
    
    Sheets("Base de dados").Activate
    On Error Resume Next
    Sheets("Base de dados").ShowAllData
    lin = Sheets("Base de dados").Cells(1, 1).End(xlDown).Row + 1
    Sheets("Base de dados").Cells(lin, 1) = lin - 1
    Sheets("Base de dados").Cells(lin, 2) = TextBox102.Text
    Sheets("Base de dados").Cells(lin, 3) = TextBox103.Text
    Sheets("Base de dados").Cells(lin, 4) = TextBox104.Text
    Sheets("Base de dados").Cells(lin, 5) = TextBox105.Text
    Sheets("Base de dados").Cells(lin, 6) = TextBox106.Text
    Sheets("Base de dados").Cells(lin, 7) = TextBox1071.Text & "/" & TextBox1072.Text
    Sheets("Base de dados").Cells(lin, 8) = TextBox108.Text
    Sheets("Base de dados").Cells(lin, 9) = TextBox109.Text
    Sheets("Base de dados").Cells(lin, 10) = TextBox110.Text
    Sheets("Base de dados").Cells(lin, 11) = Now
    
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    
    result = MsgBox("CADASTRO REALIZADO COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA LIMPAR OS DADOS?", vbYesNo + vbInformation)
    If result = vbYes Then
        TextBox101 = ""
        TextBox102 = ""
        TextBox103 = ""
        TextBox104 = ""
        TextBox105 = ""
        TextBox106 = ""
        TextBox1071 = ""
        TextBox1072 = ""
        TextBox108 = ""
        TextBox109 = ""
        TextBox110 = ""
    End If
    
    TextBox102.SetFocus
    
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub CommandButton102_Click()
'LIMPAR CADASTRO
    TextBox101 = ""
    TextBox102 = ""
    TextBox103 = ""
    TextBox104 = ""
    TextBox105 = ""
    TextBox106 = ""
    TextBox1071 = ""
    TextBox1072 = ""
    TextBox108 = ""
    TextBox109 = ""
    TextBox110 = ""
End Sub

Private Sub CommandButton103_Click()
'CANCELAR CADASTRO
    result = MsgBox("DESEJA SALVAR AS ALTERAÇÕES ?", vbYesNo + vbCritical)
    If result = vbYes Then
        ThisWorkbook.Save
    End If
    End
End Sub

Private Sub CommandButton104_Click()
'ATUALIZAR CADASTRO
    If TextBox101 = "" Then
        MsgBox "CADASTRO NÃO EXISTE!" & vbCrLf & vbCrLf & "CLIQUE EM CADASTRAR", vbCritical
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
    Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=TextBox101.Text
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    Sheets("Base de dados").Cells(lin_inicio, 2) = TextBox102.Text
    Sheets("Base de dados").Cells(lin_inicio, 3) = TextBox103.Text
    Sheets("Base de dados").Cells(lin_inicio, 4) = TextBox104.Text
    Sheets("Base de dados").Cells(lin_inicio, 5) = TextBox105.Text
    Sheets("Base de dados").Cells(lin_inicio, 6) = TextBox106.Text
    Sheets("Base de dados").Cells(lin_inicio, 7) = TextBox1071.Text & "/" & TextBox1072.Text
    Sheets("Base de dados").Cells(lin_inicio, 8) = TextBox108.Text
    Sheets("Base de dados").Cells(lin_inicio, 9) = TextBox109.Text
    Sheets("Base de dados").Cells(lin_inicio, 10) = TextBox110.Text
    Sheets("Base de dados").Cells(lin_inicio, 12) = Now
    
    Sheets("Base de dados").ShowAllData
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    
    result = MsgBox("CADASTRO ATUALIZADO COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA LIMPAR OS DADOS?", vbYesNo + vbInformation)
    If result = vbYes Then
        TextBox101 = ""
        TextBox102 = ""
        TextBox103 = ""
        TextBox104 = ""
        TextBox105 = ""
        TextBox106 = ""
        TextBox1071 = ""
        TextBox1072 = ""
        TextBox108 = ""
        TextBox109 = ""
        TextBox110 = ""
        UserForm1.MultiPage1.Value = 0
    End If
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub CommandButton201_Click()
'OK BAIXA

    If TextBox201 = "" And TextBox202 = "" And TextBox203 = "" And TextBox204 = "" Then
        MsgBox "FAVOR PREENCHER AS INFORMAÇÕES DE BAIXA!", vbCritical
        GoTo FIM
    End If
    If TextBox204 <> "" And OptionButton1.Value = False And OptionButton2.Value = False Then
        MsgBox "FAVOR PREENCHER O TIPO DE BAIXA!", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    Sheets("Base de dados").Visible = xlSheetVisible
    Sheets("Base de dados").Activate
    If Not Sheets("Base de dados").AutoFilterMode Then
        Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
    End If
    On Error Resume Next
NOVO:
    Sheets("Base de dados").ShowAllData
    Set rngAF = Range("A1:A" & Cells(1, 1).End(xlDown).Row)
    
    If TextBox201 <> "" Then
        Sheets("Base de dados").Range("D:D").AutoFilter Field:=4, Criteria1:=TextBox201.Text
    End If
    If TextBox202 <> "" Then
        Sheets("Base de dados").Range("E:E").AutoFilter Field:=5, Criteria1:=TextBox202.Text
    End If
    If TextBox203 <> "" Then
        Sheets("Base de dados").Range("C:C").AutoFilter Field:=3, Criteria1:=TextBox203.Text
    End If
    If TextBox204 <> "" Then
        Sheets("Base de dados").Range("I:I").AutoFilter Field:=9, Criteria1:=TextBox204.Text
    End If
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    If Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALOR INFORMADO NÃO EXISTE !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR INFORMAÇÕES", vbCritical
        GoTo FIM
    Else
        Dim arrayItems2()
        With Planilha5
            ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, Sheets("Base de dados").Range("A:A")), 1 To 10) '.UsedRange.Columns.Count
            Me.ListBox2.ColumnCount = 10 '.UsedRange.Columns.Count
            Me.ListBox2.ColumnWidths = "40;90;80;100;100;180;40;50;80;200"
            i = 0
            For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
                Me.ListBox2.AddItem
                For coluna = 1 To 10 '.UsedRange.Columns.Count
                    arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
                Next coluna
                i = i + 1
            Next rngcell
            Me.ListBox2.List = arrayItems2()
        End With
    End If
    
    If OptionButton1.Value = True Then
        result = MsgBox("TEM CERTEZA QUE DESEJA DAR BAIXA TODOS OS ITENS DA POSIÇÃO " & TextBox204 & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
        If result = vbYes Then
            Do While Cells(lin_inicio, 1).Value <> 0
                Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Baixa").Rows(Sheets("Baixa").Cells(1, 1).End(xlDown).Row + 1)
                Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 13) = Now
                Sheets("Base de dados").Rows(lin_inicio).Delete
                lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            Loop
            Sheets("Base de dados").ShowAllData
        
            Sheets("Base de dados").Cells(2, 1) = 1
            Sheets("Base de dados").Cells(3, 1) = 2
            Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 1).End(xlDown).Row)
        Else
            Sheets("Base de dados").ShowAllData
            GoTo FIM
        End If
        MsgBox "BAIXA REALIZADA COM SUCESSO !", vbInformation
        TextBox201 = ""
        TextBox202 = ""
        TextBox203 = ""
        TextBox204 = ""
        ListBox2.Clear
        OptionButton1.Value = False
        OptionButton2.Value = False
        OptionButton1.Enabled = False
        OptionButton2.Enabled = False
        GoTo FIM
    End If
    
    ID = Application.InputBox("INFORME O ID")
    If ID = 0 Then
        Sheets("Base de dados").ShowAllData
        Sheets("Base de dados").Cells(2, 1) = 1
        Sheets("Base de dados").Cells(3, 1) = 2
        Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 1).End(xlDown).Row)
        GoTo FIM
    End If
ERRO:
    result = Application.InputBox("INFORME A QUANTIDADE QUE DESEJA DAR BAIXA NO ID " & ID & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA")
    If result > 0 Then
        If Not Sheets("Base de dados").AutoFilterMode Then
            Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
        End If
        On Error Resume Next
        Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=ID
        lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        If Sheets("Base de dados").Cells(lin_inicio, 8).Value - result > 0 Then
            Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Baixa").Rows(Sheets("Baixa").Cells(1, 1).End(xlDown).Row + 1)
            Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 8) = result
            Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 13) = Now
            Sheets("Base de dados").Cells(lin_inicio, 8).Value = Sheets("Base de dados").Cells(lin_inicio, 8).Value - result
        ElseIf Sheets("Base de dados").Cells(lin_inicio, 8).Value - result = 0 Then
            Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Baixa").Rows(Sheets("Baixa").Cells(1, 1).End(xlDown).Row + 1)
            Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 13) = Now
            Sheets("Base de dados").Rows(lin_inicio).Delete
        Else
            MsgBox "VALOR INFORMADO MAIOR QUE ESTOQUE" & vbCrLf & vbCrLf & "VERIFIQUE A QUANTIDADE INFORMADA"
            GoTo ERRO
        End If
        
        Sheets("Base de dados").ShowAllData
        result = MsgBox("BAIXA REALIZADA COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA BAIXAR OUTRO ITEM DESSE PEDIDO ?", vbYesNo + vbInformation)
        If result = vbYes Then
            GoTo NOVO
        Else
            Sheets("Base de dados").Cells(2, 1) = 1
            Sheets("Base de dados").Cells(3, 1) = 2
            Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 1).End(xlDown).Row)
        End If
    End If

FIM:
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton202_Click()
'LIMPAR BAIXA
    TextBox201 = ""
    TextBox202 = ""
    TextBox203 = ""
    TextBox204 = ""
    ListBox2.Clear
    OptionButton1.Value = False
    OptionButton2.Value = False
    OptionButton1.Enabled = False
    OptionButton2.Enabled = False
End Sub

Private Sub CommandButton203_Click()
'CANCELAR BAIXA
    result = MsgBox("DESEJA SALVAR AS ALTERAÇÕES ?", vbYesNo + vbCritical)
    If result = vbYes Then
        ThisWorkbook.Save
    End If
    End
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If IsNumeric(Left(TextBox6, 1)) Then
        TextBox6.MaxLength = 8
        Select Case KeyAscii
            Case 8 'Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}" 'Emula o TAB
            Case 48 To 57
            If TextBox6.SelStart = 2 Then
                TextBox6.SelText = "-"
            End If
            If TextBox6.SelStart = 5 Then
                TextBox6.SelText = "-"
            End If
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
        End Select
    Else
        TextBox6.MaxLength = 7
        Select Case KeyAscii
            Case 8 'Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}" 'Emula o TAB
            Case 48 To 90
            If TextBox6.SelStart = 1 Then
                TextBox6.SelText = "-"
            End If
            If TextBox6.SelStart = 4 Then
                TextBox6.SelText = "-"
            End If
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
        End Select
    End If
End Sub

Private Sub TextBox103_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox104_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox105_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox109_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If IsNumeric(Left(TextBox109, 1)) Then
        TextBox109.MaxLength = 8
        Select Case KeyAscii
            Case 8 'Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}" 'Emula o TAB
            Case 48 To 57
            If TextBox109.SelStart = 2 Then
                TextBox109.SelText = "-"
            End If
            If TextBox109.SelStart = 5 Then
                TextBox109.SelText = "-"
            End If
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
        End Select
    Else
        TextBox109.MaxLength = 7
        Select Case KeyAscii
            Case 8 'Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}" 'Emula o TAB
            Case 48 To 90
            If TextBox109.SelStart = 1 Then
                TextBox109.SelText = "-"
            End If
            If TextBox109.SelStart = 4 Then
                TextBox109.SelText = "-"
            End If
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
        End Select
    End If
End Sub

Private Sub TextBox110_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox203_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox104_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If TextBox104 <> "" Then
        TextBox106 = Application.WorksheetFunction.VLookup(TextBox104, Sheets("Descrição").Range("A:B"), 2, 0)
    Else
        TextBox106 = ""
    End If
End Sub

Private Sub TextBox204_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'KeyAscii = Asc(UCase(Chr(KeyAscii)))
    TextBox204.MaxLength = 8
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        If TextBox204.SelStart = 2 Then
            TextBox204.SelText = "-"
        End If
        If TextBox204.SelStart = 5 Then
            TextBox204.SelText = "-"
        End If
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox204_Change()
    If TextBox204 <> "" And Len(TextBox204) = 8 Then
        OptionButton1.Enabled = True
        OptionButton2.Enabled = True
    Else
        OptionButton1.Enabled = False
        OptionButton2.Enabled = False
        OptionButton1.Value = False
        OptionButton2.Value = False
    End If
End Sub
