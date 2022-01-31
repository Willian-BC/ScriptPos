Private Sub lsDesconectar()
    If gConexao.State = adStateOpen Then
        gConexao.Close
        Set gConexao = Nothing
        MsgBox " Conexão fechada "
    End If
End Sub
