Public Sub lsExcluirDados()
    Set ws = Sheets("Excluir")
    tabela = "BD_dados"
    
    nID = CInt(ws.Cells(2, 1))
    
    lsConectar
    lrs.Open "SELECT * FROM BD_dados WHERE ID = " & nID, gConexao, adOpenKeyset, adLockPessimistic
    
    If lrs.RecordCount = 1 Then
        lrs.Delete
        lrs.Update
        MsgBox " Excluido com sucesso ", vbExclamation
    Else
        MsgBox " ID não encontrado ", vbExclamation
    End If
    
    lsDesconectar
End Sub
