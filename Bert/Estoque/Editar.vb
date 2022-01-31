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
    Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=ID
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    If Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR O NÚMERO DE ID", vbCritical
        Sheets("Base de dados").Range("A:A").AutoFilter Field:=1
        GoTo FIM
    End If
    
    TextBox109 = Sheets("Base de dados").Cells(lin_inicio, 1)
    TextBox101 = Sheets("Base de dados").Cells(lin_inicio, 2)
    TextBox111 = Sheets("Base de dados").Cells(lin_inicio, 3)
    TextBox102 = Sheets("Base de dados").Cells(lin_inicio, 4)
    TextBox103 = Sheets("Base de dados").Cells(lin_inicio, 5)
    TextBox104 = Sheets("Base de dados").Cells(lin_inicio, 6)
    TextBox105 = Sheets("Base de dados").Cells(lin_inicio, 7)
    TextBox106 = Sheets("Base de dados").Cells(lin_inicio, 8)
    TextBox107 = Sheets("Base de dados").Cells(lin_inicio, 9)
    TextBox108 = Sheets("Base de dados").Cells(lin_inicio, 10)
    TextBox110 = Sheets("Base de dados").Cells(lin_inicio, 11)
    
    Sheets("Base de dados").ShowAllData
    
    ListBox1.Clear
    UserForm1.MultiPage1.Value = 1
FIM:
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
End Sub
