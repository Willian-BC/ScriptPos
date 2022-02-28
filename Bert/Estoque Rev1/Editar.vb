Private Sub CommandButton4_Click()
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Set ws = Sheets("Base de dados")
ERRO:
    ID = Application.InputBox("INFORME O ID")
    If ID = 0 Then GoTo FIM
    
    If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    ws.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    If ws.Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR O NÚMERO DE ID", vbCritical
        ws.Range("A:A").AutoFilter Field:=1
        GoTo ERRO
    End If
    
    TextBox101 = ws.Cells(lin_inicio, 1)
    TextBox102 = ws.Cells(lin_inicio, 5)
    TextBox103 = ws.Cells(lin_inicio, 2)
    TextBox104 = ws.Cells(lin_inicio, 3)
    TextBox105 = ws.Cells(lin_inicio, 4)
    TextBox106 = ws.Cells(lin_inicio, 6)
    TextBox107 = ws.Cells(lin_inicio, 7)
    TextBox108 = ws.Cells(lin_inicio, 8)
    TextBox109 = ws.Cells(lin_inicio, 9)
    TextBox110 = ws.Cells(lin_inicio, 10)
    TextBox111 = ws.Cells(lin_inicio, 11)
    
    ws.ShowAllData
    ListBox1.Clear
    UserForm1.MultiPage1.Value = 1
    Application.ScreenUpdating = True
FIM:
End Sub
