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
    Sheets("Base de dados").Cells(lin_inicio, 3) = TextBox111.Text
    Sheets("Base de dados").Cells(lin_inicio, 4) = TextBox102.Text
    Sheets("Base de dados").Cells(lin_inicio, 5) = TextBox103.Text
    Sheets("Base de dados").Cells(lin_inicio, 6) = TextBox104.Text
    Sheets("Base de dados").Cells(lin_inicio, 7) = TextBox105.Text
    Sheets("Base de dados").Cells(lin_inicio, 8) = TextBox106.Text
    Sheets("Base de dados").Cells(lin_inicio, 9) = TextBox107.Text
    Sheets("Base de dados").Cells(lin_inicio, 10) = TextBox108.Text
    Sheets("Base de dados").Cells(lin_inicio, 11) = TextBox110.Text
    Sheets("Base de dados").Cells(lin_inicio, 13) = Now
    
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
        TextBox111 = ""
        UserForm1.MultiPage1.Value = 0
    End If
    Application.ScreenUpdating = True
FIM:
End Sub
