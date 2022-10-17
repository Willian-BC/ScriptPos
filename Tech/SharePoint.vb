'Tools > References > Active "Microsoft ActiveX Data Objects 6.1 Library"
'Settings > List Settings > Copy URL "List=%7B88C226AA-F970-413D-8EB6-7291B52C82D5%7D"
Dim cnt As ADODB.Connection
Dim rst As ADODB.Recordset
Dim mySQL As String

Private Sub lsConectar()
    Set cnt = New ADODB.Connection
    With cnt
        .ConnectionString = _
        "Provider = Microsoft.ACE.OLEDB.12.0;WSS;IMEX=0;RetrieveIds=Yes;DATABASE=https://digicorner.sharepoint.com/sites/ControleEstatistico;LIST={88C226AA-F970-413D-8EB6-7291B52C82D5};"
        .Open
    End With
End Sub

Private Sub lsDesconectar()
    If CBool(rst.State And adStateOpen) = True Then rst.Close
    Set rst = Nothing
    If CBool(cnt.State And adStateOpen) = True Then cnt.Close
    Set cnt = Nothing
End Sub

Private Sub CommandButton1_Click()
'OK CADASTRO'
    Set rst = New ADODB.Recordset
    
    lsConectar
    mySQL = "SELECT * FROM [DataBase];"
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
    
    lsDesconectar
End Sub

Private Sub CommandButton2_Click()
'OK PESQUISA
    Set rst = New ADODB.Recordset
    Set ws = Sheets("Base de dados")
    
    ws.Rows("2:" & ws.Cells(2, 1).End(xlDown).Row).ClearContents
    lsConectar
    mySQL = "SELECT * FROM [DataBase];"
    rst.Open mySQL, cnt, adOpenDynamic, adLockOptimistic
    
    ws.Cells(2, 1).CopyFromRecordset rst
    
    lsDesconectar
End Sub

Private Sub CommandButton3_Click()
'OK DELETAR
    Set rst = New ADODB.Recordset
    
    ID = TextBox0
    lsConectar
    mySQL = "DELETE * FROM [DataBase] WHERE [ID] = " & CInt(ID) & ";"
    cnt.Execute mySQL, , adCmdText
    
    lsDesconectar
End Sub

Private Sub CommandButton4_Click()
'OK ATUALIZAR
    Set rst = New ADODB.Recordset
    
    ID = TextBox0
    lsConectar
    mySQL = "SELECT * FROM [DataBase] WHERE [ID] = " & CInt(ID) & ";"
    rst.Open mySQL, cnt, adOpenDynamic, adLockOptimistic
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

    lsDesconectar
End Sub
Public Sub UserForm_Initialize()
    ComboBox1.AddItem "Homologação"
    ComboBox1.AddItem "Produção"
    
    ComboBox2.AddItem "MAQ-01"
    ComboBox2.AddItem "MAQ-02"
    ComboBox2.AddItem "MAQ-03"
    ComboBox2.AddItem "MAQ-04"
    ComboBox2.AddItem "MAQ-05"
    ComboBox2.AddItem "MAQ-06"
    
    ComboBox3.AddItem "MAQ-01"
    ComboBox3.AddItem "MAQ-02"
    ComboBox3.AddItem "MAQ-03"
    ComboBox3.AddItem "MAQ-04"
    ComboBox3.AddItem "MAQ-05"
    ComboBox3.AddItem "MAQ-06"
End Sub
