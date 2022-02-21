Private Sub UserForm_Initialize()
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    Sheets("Baixa").Visible = xlSheetVeryHidden
    Sheets("Excluido").Visible = xlSheetVeryHidden
End Sub
Private Sub CommandButton100_Click()
    Dim MyValue As Variant
    MyValue = InputBox("Digite a senha")
    If MyValue = "1010" Then
        Sheets("Base de dados").Visible = xlSheetVisible
        Sheets("Baixa").Visible = xlSheetVisible
        Sheets("Excluido").Visible = xlSheetVisible
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
    
    If TextBox1 <> "" Then
        Sheets("Base de dados").Range("B:B").AutoFilter Field:=2, Criteria1:=TextBox1.Text
    End If
    If TextBox2 <> "" Then
        Sheets("Base de dados").Range("C:C").AutoFilter Field:=3, Criteria1:=TextBox2.Text
    End If
    If TextBox3 <> "" Then
        Sheets("Base de dados").Range("D:D").AutoFilter Field:=4, Criteria1:=TextBox3.Text
    End If
    If TextBox4 <> "" Then
        Sheets("Base de dados").Range("E:E").AutoFilter Field:=5, Criteria1:=TextBox4.Text
    End If
    If TextBox5 <> "" Then
        Sheets("Base de dados").Range("F:F").AutoFilter Field:=6, Criteria1:="=*" & TextBox5.Text & "*"
    End If
    If TextBox6 <> "" Then
        Sheets("Base de dados").Range("H:H").AutoFilter Field:=8, Criteria1:=TextBox6.Text
    End If
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = Sheets("Base de dados").Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If Cells(lin_inicio, 1).Value = 0 Then
            MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
            Sheets("Base de dados").ShowAllData
            GoTo FIM
    Else
    
    Dim arrayItems2()
        With Planilha5
            ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, Sheets("Base de dados").Range("C:C")), 1 To 10) '.UsedRange.Columns.Count
            Me.ListBox1.ColumnCount = 10 '.UsedRange.Columns.Count
            Me.ListBox1.ColumnWidths = "30;130;80;90;80;90;90;80;80;200"
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
        Set rngAJ = Range("A1:K" & lin_fim).SpecialCells(xlCellTypeVisible)
        rngAJ.Copy
        Workbooks.Add
        Range("A1").PasteSpecial Paste:=xlPasteValues
        Columns("K:K").NumberFormat = "dd/mm/yyyy"
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
    
    TextBox109 = Sheets("Base de dados").Cells(lin_inicio, 1)
    TextBox101 = Sheets("Base de dados").Cells(lin_inicio, 2)
    TextBox102 = Sheets("Base de dados").Cells(lin_inicio, 3)
    TextBox103 = Sheets("Base de dados").Cells(lin_inicio, 4)
    TextBox104 = Sheets("Base de dados").Cells(lin_inicio, 5)
    TextBox105 = Sheets("Base de dados").Cells(lin_inicio, 6)
    TextBox106 = Sheets("Base de dados").Cells(lin_inicio, 7)
    TextBox107 = Sheets("Base de dados").Cells(lin_inicio, 8)
    TextBox108 = Sheets("Base de dados").Cells(lin_inicio, 9)
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
            Sheets("Excluido").Cells(Sheets("Excluido").Cells(1, 1).End(xlDown).Row, 13) = Date
            Rows(lin_inicio).Delete
            
            Sheets("Base de dados").ShowAllData
            
            Sheets("Base de dados").Cells(2, 1) = 1
            Sheets("Base de dados").Cells(3, 1) = 2
            Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 2).End(xlDown).Row)
            
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
    If TextBox109 <> "" Then
        MsgBox "CADASTRO JÁ EXISTE!" & vbCrLf & vbCrLf & "CLIQUE EM ATUALIZAR REGISTRO OU LIMPAR", vbCritical
        GoTo FIM
    ElseIf TextBox101 = "" Or TextBox102 = "" Or TextBox103 = "" Or TextBox104 = "" Or TextBox105 = "" Or TextBox106 = "" Or TextBox107 = "" Or TextBox108 = "" Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS PARA CADASTRO!", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    Sheets("Base de dados").Visible = xlSheetVisible
    
    Sheets("Base de dados").Activate
    On Error Resume Next
    Sheets("Base de dados").ShowAllData
    lin = Sheets("Base de dados").Cells(1, 1).End(xlDown).Row + 1
    Sheets("Base de dados").Cells(lin, 1) = Sheets("Base de dados").Cells(lin - 1, 1) + 1
    Sheets("Base de dados").Cells(lin, 2) = TextBox101.Text
    Sheets("Base de dados").Cells(lin, 3) = TextBox102.Text
    Sheets("Base de dados").Cells(lin, 4) = TextBox103.Text
    Sheets("Base de dados").Cells(lin, 5) = TextBox104.Text
    Sheets("Base de dados").Cells(lin, 6) = TextBox105.Text
    Sheets("Base de dados").Cells(lin, 7) = TextBox106.Text
    Sheets("Base de dados").Cells(lin, 8) = TextBox107.Text
    Sheets("Base de dados").Cells(lin, 9) = TextBox108.Text
    Sheets("Base de dados").Cells(lin, 10) = TextBox110.Text
    Sheets("Base de dados").Cells(lin, 11) = Date
    
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    
    result = MsgBox("CADASTRO REALIZADO COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA LIMPAR OS DADOS?", vbYesNo + vbInformation)
    If result = vbYes Then
        TextBox101 = ""
        TextBox102 = ""
        TextBox103 = ""
        TextBox104 = ""
        TextBox105 = ""
        TextBox106 = ""
        TextBox107 = ""
        TextBox108 = ""
        TextBox109 = ""
        TextBox110 = ""
    End If
    
    TextBox101.SetFocus
    
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
    TextBox107 = ""
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
    If TextBox109 = "" Then
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
    Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=TextBox109.Text
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    Sheets("Base de dados").Cells(lin_inicio, 2) = TextBox101.Text
    Sheets("Base de dados").Cells(lin_inicio, 3) = TextBox102.Text
    Sheets("Base de dados").Cells(lin_inicio, 4) = TextBox103.Text
    Sheets("Base de dados").Cells(lin_inicio, 5) = TextBox104.Text
    Sheets("Base de dados").Cells(lin_inicio, 6) = TextBox105.Text
    Sheets("Base de dados").Cells(lin_inicio, 7) = TextBox106.Text
    Sheets("Base de dados").Cells(lin_inicio, 8) = TextBox107.Text
    Sheets("Base de dados").Cells(lin_inicio, 9) = TextBox108.Text
    Sheets("Base de dados").Cells(lin_inicio, 10) = TextBox110.Text
    Sheets("Base de dados").Cells(lin_inicio, 12) = Date
    
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
        TextBox107 = ""
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

    If TextBox201 = "" Or TextBox202 = "" Or (OptionButton1.Value = False And OptionButton2.Value = False) Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS!", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    Sheets("Base de dados").Visible = xlSheetVisible
    Sheets("Base de dados").Activate
NOVO:
    If Not Sheets("Base de dados").AutoFilterMode Then
        Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
    End If
    On Error Resume Next
    Sheets("Base de dados").ShowAllData
    
    Set rngAF = Range("A1:A" & Cells(1, 1).End(xlDown).Row)
    
    If TextBox201 <> "" Then
        Sheets("Base de dados").Range("C:C").AutoFilter Field:=3, Criteria1:=TextBox201.Text
    End If
    If TextBox202 <> "" Then
        Sheets("Base de dados").Range("F:F").AutoFilter Field:=6, Criteria1:="=*" & TextBox202.Text & "*"
    End If
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    If Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "PEDIDO INFORMADO NÃO EXISTE !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR O CÓD E NOME DO CLIENTE", vbCritical
        GoTo FIM
    Else
        Dim arrayItems2()
            With Planilha5
                ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, Sheets("Base de dados").Range("C:C")), 1 To 10) '.UsedRange.Columns.Count
                Me.ListBox2.ColumnCount = 10 '.UsedRange.Columns.Count
                Me.ListBox2.ColumnWidths = "30;130;80;90;80;90;90;80;80;200"
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
'        If Not Sheets("Base de dados").AutoFilterMode Then
'            Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
'        End If
'        On Error Resume Next
'        Sheets("Base de dados").ShowAllData
'
'        If TextBox201 <> "" Then
'            Sheets("Base de dados").Range("C:C").AutoFilter Field:=3, Criteria1:=TextBox201.Text
'        End If
'        If TextBox202 <> "" Then
'            Sheets("Base de dados").Range("F:F").AutoFilter Field:=6, Criteria1:="=*" & TextBox202.Text & "*"
'        End If
'
'        lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
'        lin_fim = Sheets("Base de dados").Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
        
        result = MsgBox("TEM CERTEZA QUE DESEJA DAR BAIXA EM TODO O PEDIDO " & TextBox201 & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
        If result = vbYes Then
            Do While Cells(lin_inicio, 1).Value <> 0
                Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Baixa").Rows(Sheets("Baixa").Cells(1, 1).End(xlDown).Row + 1)
                Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 13) = Date
                Sheets("Base de dados").Rows(lin_inicio).Delete
                lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            Loop
            Sheets("Base de dados").ShowAllData
        
            Sheets("Base de dados").Cells(2, 1) = 1
            Sheets("Base de dados").Cells(3, 1) = 2
            Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 2).End(xlDown).Row)
        Else
            GoTo FIM
        End If
'        Sheets("Base de dados").ShowAllData
        MsgBox "BAIXA REALIZADA COM SUCESSO !", vbInformation
        TextBox201 = ""
        TextBox202 = ""
        OptionButton1.Value = False
        OptionButton2.Value = False
        ListBox2.Clear
        GoTo FIM
    ElseIf OptionButton2.Value = True Then
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
            If Not Sheets("Base de dados").AutoFilterMode Then
                Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
            End If
            On Error Resume Next
            Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=ID
            lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            
            Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Baixa").Rows(Sheets("Baixa").Cells(1, 1).End(xlDown).Row + 1)
            Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 13) = Date
            Sheets("Base de dados").Rows(lin_inicio).Delete
            
            Sheets("Base de dados").ShowAllData
            result = MsgBox("BAIXA REALIZADA COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA BAIXAR OUTRO ITEM DESSE PEDIDO ?", vbYesNo + vbInformation)
            If result = vbYes Then
                GoTo NOVO
            Else
                Sheets("Base de dados").Cells(2, 1) = 1
                Sheets("Base de dados").Cells(3, 1) = 2
                Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 2).End(xlDown).Row)
            End If
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
    OptionButton1.Value = False
    OptionButton2.Value = False
    ListBox2.Clear
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

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox101_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox104_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox105_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox107_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox108_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox110_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox202_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

'Private Sub txtdata_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'
'    txtdata.MaxLength = 10
'    Select Case KeyAscii
'        Case 8 'Aceita o BACK SPACE
'        Case 13: SendKeys "{TAB}" 'Emula o TAB
'        Case 48 To 57
'        If txtdata.SelStart = 2 Then
'            txtdata.SelText = "/"
'        End If
'        If txtdata.SelStart = 5 Then
'            txtdata.SelText = "/"
'        End If
'    Case Else: KeyAscii = 0 'Ignora os outros caracteres
'    End Select
'
'
'End Sub
'
'Private Sub txthora_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'    'Excel Flex - www.excelflex.com.br/dicas
'    If Not IsNumeric(Chr(KeyAscii.Value)) Or Len(txthora.Text) >= 5 Then
'        KeyAscii.Value = 0
'    Else
'        If Len(txthora.Text) = 2 Then
'            txthora.Text = txthora.Text & ":"
'        End If
'    End If
'
'
'End Sub
'
'Private Sub txtsolicitante_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'
'KeyAscii = Asc(UCase(Chr(KeyAscii)))
'
'End Sub
