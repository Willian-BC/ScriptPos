Public Sub lsListarDados()
    Set lrs = New ADODB.recordset
    tabela = "BD_dados"
    Set ws = Sheets("BD_dados")
    
    lsConectar
    lrs.Open "Select * from " & tabela, gConexao
    
    ws.Columns("A:M").ClearContents
    ws.Cells(2, 1).CopyFromRecordset lrs
    
    If Not lrs Is Nothing Then
        lrs.Close
        Set lrs = Nothing
    End If
    lsDesconectar
End Sub
