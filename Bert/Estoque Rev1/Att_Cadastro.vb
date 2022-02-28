Private Sub CommandButton104_Click()
    If TextBox101 = "" Then
        MsgBox "CADASTRO NÃO EXISTE!" & vbCrLf & vbCrLf & "CLIQUE EM CADASTRAR", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    
    lsConectar
    Set lrs = New ADODB.Recordset
    lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(TextBox101), gConexao, adOpenKeyset, adLockPessimistic
    
    lrs!Item = TextBox103
    lrs!descricao = TextBox104
    lrs!pedido = TextBox105
    lrs!ordem_prod = TextBox102
    lrs!programa = TextBox106
    lrs!cliente = TextBox107
    lrs!qt_pecas = TextBox108
    lrs!area_estoque = TextBox109
    lrs!posicao = TextBox110
    lrs!comentario = TextBox111
    lrs!edicao = Now
    
    lrs.Update
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
    
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
