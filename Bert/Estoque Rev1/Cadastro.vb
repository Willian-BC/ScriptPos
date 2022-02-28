Private Sub CommandButton101_Click()
    If TextBox101 <> "" Then
        MsgBox "CADASTRO JÁ EXISTE!" & vbCrLf & vbCrLf & "CLIQUE EM ATUALIZAR REGISTRO OU LIMPAR", vbCritical
        GoTo FIM
    ElseIf TextBox102 = "" Or TextBox103 = "" Or TextBox105 = "" Or TextBox108 = "" Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS PARA CADASTRO!", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    
    lsConectar
    Set lrs = New ADODB.Recordset
    
    sql = " INSERT INTO BD_dados "
    sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada) "
    sql = sql & " VALUES "
    sql = sql & " ('" & TextBox103 & "', "
    sql = sql & " '" & TextBox104 & "', "
    sql = sql & " '" & TextBox105 & "', "
    sql = sql & " '" & TextBox102 & "', "
    sql = sql & " '" & TextBox106 & "', "
    sql = sql & " '" & TextBox107 & "', "
    sql = sql & " '" & TextBox108 & "', "
    sql = sql & " '" & TextBox109 & "', "
    sql = sql & " '" & TextBox110 & "', "
    sql = sql & " '" & TextBox111 & "', "
    sql = sql & " '" & Now & "') "
    
    lrs.Open sql, gConexao
    Set lrs = Nothing
    lsDesconectar
    
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
        TextBox102.SetFocus
    End If
    
    Application.ScreenUpdating = True
FIM:
End Sub
