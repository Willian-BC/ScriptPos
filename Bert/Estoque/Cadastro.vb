Private Sub CommandButton101_Click()
'OK CADASTRO'
    If TextBox109 <> "" Then
        MsgBox "CADASTRO JÁ EXISTE!" & vbCrLf & vbCrLf & "CLIQUE EM ATUALIZAR REGISTRO OU LIMPAR", vbCritical
        GoTo FIM
    ElseIf TextBox101 = "" Or TextBox102 = "" Or TextBox103 = "" Or TextBox105 = "" Or TextBox106 = "" Then
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
    Sheets("Base de dados").Cells(lin, 3) = TextBox111.Text
    Sheets("Base de dados").Cells(lin, 4) = TextBox102.Text
    Sheets("Base de dados").Cells(lin, 5) = TextBox103.Text
    Sheets("Base de dados").Cells(lin, 6) = TextBox104.Text
    Sheets("Base de dados").Cells(lin, 7) = TextBox105.Text
    Sheets("Base de dados").Cells(lin, 8) = TextBox106.Text
    Sheets("Base de dados").Cells(lin, 9) = TextBox107.Text
    Sheets("Base de dados").Cells(lin, 10) = TextBox108.Text
    Sheets("Base de dados").Cells(lin, 11) = TextBox110.Text
    Sheets("Base de dados").Cells(lin, 12) = Now
    
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
        TextBox111 = ""
    End If
    
    TextBox103.SetFocus
    ThisWorkbook.Save
    
    Application.ScreenUpdating = True
FIM:
End Sub
