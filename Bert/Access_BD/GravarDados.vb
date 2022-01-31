Public Sub lsGravarDados()
    Set ws = Sheets("Inserir")
    tabela = "BD_dados"
    
    lsConectar
    sql = " INSERT INTO " & tabela
    sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada, edicao) "
    sql = sql & " VALUES "
    sql = sql & " ('" & ws.Cells(2, 1) & "', "
    sql = sql & " '" & ws.Cells(2, 2) & "', "
    sql = sql & " '" & ws.Cells(2, 3) & "', "
    sql = sql & " '" & ws.Cells(2, 4) & "', "
    sql = sql & " '" & ws.Cells(2, 5) & "', "
    sql = sql & " '" & ws.Cells(2, 6) & "', "
    sql = sql & " '" & ws.Cells(2, 7) & "', "
    sql = sql & " '" & ws.Cells(2, 8) & "', "
    sql = sql & " '" & ws.Cells(2, 9) & "', "
    sql = sql & " '" & ws.Cells(2, 10) & "', "
    sql = sql & " '" & ws.Cells(2, 11) & "', "
    sql = sql & " '" & ws.Cells(2, 12) & "') "
    
    lrs.Open sql, gConexao
    
    lsDesconectar
    MsgBox " Inclusão realizada com sucesso "
End Sub
