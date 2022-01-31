Private Sub CommandButton100_Click()
    Dim MyValue As Variant
    MyValue = InputBox("Digite a senha")
    If MyValue = "VALOR" Then
        Sheets("Base de dados").Visible = xlSheetVisible
        Sheets("Baixa").Visible = xlSheetVisible
        Sheets("Excluido").Visible = xlSheetVisible
        Sheets("PRJ").Visible = xlSheetVisible
        Sheets("Componentes").Visible = xlSheetVisible
    Else
        MsgBox ("Senha Incorreta")
    End If
End Sub
