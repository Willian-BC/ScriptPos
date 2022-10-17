'Tools > References > Active "Microsoft ActiveX Data Objects 6.1 Library"

Private Sub CommandButton1_Click()
'OK CADASTRO'

Dim cnt As ADODB.Connection
Dim rst As ADODB.Recordset
Dim mySQL As String

Set cnt = New ADODB.Connection
Set rst = New ADODB.Recordset

    mySQL = "SELECT * FROM [Produtos];"
    
    With cnt
        .ConnectionString = _
        "Provider = Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=https://digicorner.sharepoint.com/sites/ControleEstatistico;LIST={8XXXX6AA-F970-413D-8EB6-729XXXXXX2D5};"
        .Open
    End With
    
    rst.Open mySQL, cnt, adOpenDynamic, adLockOptimistic
    
    rst.AddNew
        rst(1).Value = TextBox1 & " (" & TextBox4 & "m)" 'CHAVE PRIMARIA = TRAMO + COTA
        rst(2).Value = TextBox1 'TRAMO
        rst(3).Value = TextBox2 'CLIENTE
        rst(4).Value = TextBox4 'COTA
        rst(5).Value = TextBox5 'TERMINACAO
        rst(6).Value = ComboBox2 'MAQUINA PRINCIPAL
        rst(7).Value = TextBox6
        rst(8).Value = TextBox7
        rst(9).Value = TextBox8
     rst.Update
    
    If CBool(rst.State And adStateOpen) = True Then rst.Close
    Set rst = Nothing
    If CBool(cnt.State And adStateOpen) = True Then cnt.Close
    Set cnt = Nothing
    
End Sub
