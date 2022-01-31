Public Sub lsEditarDados()
    Set ws = Sheets("Editar")
    tabela = "BD_dados"
    
    nID = CInt(ws.Cells(2, 1))
    
    lsConectar
    lrs.Open "SELECT * FROM BD_dados WHERE ID = " & nID, gConexao, adOpenKeyset, adLockPessimistic
    
    If lrs.RecordCount = 1 Then
        lrs!item = ws.Cells(2, 2)
        lrs!descricao = ws.Cells(2, 3)
        lrs!pedido = ws.Cells(2, 4)
        lrs!ordem_prod = ws.Cells(2, 5)
        lrs!programa = ws.Cells(2, 6)
        lrs!cliente = ws.Cells(2, 7)
        lrs!qt_pecas = ws.Cells(2, 8)
        lrs!area_estoque = ws.Cells(2, 9)
        lrs!posicao = ws.Cells(2, 10)
        lrs!comentario = ws.Cells(2, 11)
        lrs!entrada = ws.Cells(2, 12)
        lrs!edicao = ws.Cells(2, 13)
        
        lrs.Update
        MsgBox " Atualização realizada com sucesso "
    Else
        MsgBox " ID não encontrado ", vbExclamation
    End If
    
    lsDesconectar
    
End Sub
