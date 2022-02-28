Private Sub CommandButton201_Click()
'OK BAIXA

    If OptionButton3.Value = True And TextBox201 = 0 And TextBox203 <> "" Then
    ElseIf TextBox201 <> "" And TextBox202 <> "" And (OptionButton1.Value = True Or OptionButton2.Value = True) Then
    ElseIf OptionButton3.Value = True And TextBox201 <> 0 Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "PEDIDO INFORMADO NÃO É KANBAN", vbCritical
        GoTo FIM
    Else
        MsgBox "FAVOR PREENCHER AS INFORMAÇÕES PARA CONTINUAR!", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Set ws = Sheets("Base de dados")
'    Set wsB = Sheets("Baixados")
    ListBox1.Clear
    
NOVO:
    If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    ws.ShowAllData
    ws.Rows("2:" & ws.Cells(1, 1).End(xlDown).Row).ClearContents
    
    lsConectar
    Set lrs = New ADODB.Recordset
    lrs.Open " SELECT * FROM BD_dados ", gConexao, adOpenKeyset, adLockPessimistic
    ws.Cells(2, 1).CopyFromRecordset lrs
    ws.Columns("A:A").NumberFormat = "0"
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
    
    Set rngAF = ws.Range("A1:A" & ws.Cells(1, 1).End(xlDown).Row)
    
    If TextBox201 <> "" Then ws.Range("D:D").AutoFilter Field:=4, Criteria1:=TextBox201
    If TextBox202 <> "" Then ws.Range("G:G").AutoFilter Field:=7, Criteria1:="=*" & TextBox202 & "*"
    If TextBox203 <> "" Then ws.Range("C:C").AutoFilter Field:=3, Criteria1:="=*" & TextBox203 & "*"
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    If ws.Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "PEDIDO INFORMADO NÃO EXISTE !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        ws.ShowAllData
        GoTo FIM
    End If

    Dim arrayItems2()
    With Planilha5
        ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, ws.Range("A:A")), 1 To 11)
        Me.ListBox2.ColumnCount = 11
        Me.ListBox2.ColumnWidths = "40;120;300;70;70;70;200;70;70;70;200"
        i = 0
        For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
            Me.ListBox2.AddItem
            For coluna = 1 To 11
                arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
            Next coluna
            i = i + 1
        Next rngcell
        Me.ListBox2.List = arrayItems2()
    End With
    
    If OptionButton1.Value = True Then
        result = MsgBox("TEM CERTEZA QUE DESEJA DAR BAIXA EM TODO O PEDIDO " & TextBox201 & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
        If result = vbYes Then
            lsConectar
            Do While ws.Cells(lin_inicio, 1).Value <> 0
                ID = ws.Cells(lin_inicio, 1)
'                Set lrs = New ADODB.Recordset
'                lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
'                wsB.Rows("2:2").ClearContents
'                wsB.Cells(2, 1).CopyFromRecordset lrs
'                wsB.Cells(2, 14) = Now
'                Set lrs = Nothing
                
                Set lrs = New ADODB.Recordset
                lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
                lrs.Delete
                lrs.Update
                Set lrs = Nothing
                
                Set lrs = New ADODB.Recordset
                sql = " INSERT INTO Baixa "
                sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada, edicao, saida) "
                sql = sql & " VALUES "
                sql = sql & " ('" & ws.Cells(lin_inicio, 2) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 3) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 4) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 5) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 6) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 7) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 8) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 9) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 10) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 11) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 12) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 13) & "', "
                sql = sql & " '" & Now & "') "
                
                lrs.Open sql, gConexao
                lrs.Close
                Set lrs = Nothing
                
                ws.Rows(lin_inicio).Delete
                lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            Loop
            lsDesconectar
        Else
            ws.ShowAllData
            GoTo FIM
        End If
        ws.ShowAllData
        MsgBox "BAIXA REALIZADA COM SUCESSO !", vbInformation
        TextBox201 = ""
        TextBox202 = ""
        TextBox203 = ""
        OptionButton1.Value = False
        OptionButton2.Value = False
        OptionButton3.Value = False
        ListBox2.Clear
        GoTo FIM
    ElseIf OptionButton2.Value = True Then
AUX:
        ID = Application.InputBox("INFORME O ID")
        If ID = 0 Then GoTo FIM
        
        result = MsgBox("TEM CERTEZA QUE DESEJA DAR BAIXA NO ID " & ID & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
        If result = vbYes Then
        
            If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
            On Error Resume Next
            ws.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
            
            lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            
            If ws.Cells(lin_inicio, 1).Value = 0 Then
                MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR O NÚMERO DE ID", vbCritical
                ws.Range("A:A").AutoFilter Field:=1
                GoTo AUX
            End If
            
            lsConectar
'            Set lrs = New ADODB.Recordset
'            lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
'            wsB.Rows("2:2").ClearContents
'            wsB.Cells(2, 1).CopyFromRecordset lrs
'            wsB.Cells(2, 14) = Now
'            Set lrs = Nothing
            
            Set lrs = New ADODB.Recordset
            lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
            lrs.Delete
            lrs.Update
            Set lrs = Nothing
            
            Set lrs = New ADODB.Recordset
            sql = " INSERT INTO Baixa "
            sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada, edicao, saida) "
            sql = sql & " VALUES "
            sql = sql & " ('" & ws.Cells(lin_inicio, 2) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 3) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 4) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 5) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 6) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 7) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 8) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 9) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 10) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 11) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 12) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 13) & "', "
            sql = sql & " '" & Now & "') "
            
            lrs.Open sql, gConexao
            lrs.Close
            Set lrs = Nothing
            lsDesconectar
            
            result = MsgBox("BAIXA REALIZADA COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA BAIXAR OUTRO ITEM DESSE PEDIDO ?", vbYesNo + vbInformation)
            If result = vbYes Then
                GoTo NOVO
            Else
                TextBox201 = ""
                TextBox202 = ""
                TextBox203 = ""
                OptionButton1.Value = False
                OptionButton2.Value = False
                OptionButton3.Value = False
                ListBox1.Clear
                ListBox2.Clear
            End If
        Else
            GoTo FIM
        End If
    ElseIf OptionButton3.Value = True Then
AUX2:
        ID = Application.InputBox("INFORME O ID")
        If ID = 0 Then GoTo FIM
        ws.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
        lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        If ws.Cells(lin_inicio, 1).Value = 0 Then
            MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR O NÚMERO DE ID", vbCritical
            ws.Range("A:A").AutoFilter Field:=1
            GoTo AUX2
        End If
ERRO:
        QTD = Application.InputBox("INFORME A QUANTIDADE QUE DESEJA DAR BAIXA NO ID " & ID & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA")
        If QTD > 0 Then
            If ws.Cells(lin_inicio, 8).Value - QTD > 0 Then
                
                lsConectar
'                Set lrs = New ADODB.Recordset
'                lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
'                wsB.Rows("2:2").ClearContents
'                wsB.Cells(2, 1).CopyFromRecordset lrs
'                wsB.Cells(2, 8) = QTD
'                wsB.Cells(2, 14) = Now
'                Set lrs = Nothing
                
                Set lrs = New ADODB.Recordset
                lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
                lrs!qt_pecas = ws.Cells(lin_inicio, 8).Value - QTD
                lrs.Update
                Set lrs = Nothing
                
                Set lrs = New ADODB.Recordset
                sql = " INSERT INTO Baixa "
                sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada, edicao, saida) "
                sql = sql & " VALUES "
                sql = sql & " ('" & ws.Cells(lin_inicio, 2) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 3) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 4) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 5) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 6) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 7) & "', "
                sql = sql & " '" & QTD & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 9) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 10) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 11) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 12) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 13) & "', "
                sql = sql & " '" & Now & "') "
                
                lrs.Open sql, gConexao
                lrs.Close
                Set lrs = Nothing
                lsDesconectar
                
            ElseIf ws.Cells(lin_inicio, 8).Value - QTD = 0 Then
                
                lsConectar
'                Set lrs = New ADODB.Recordset
'                lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
'                wsB.Rows("2:2").ClearContents
'                wsB.Cells(2, 1).CopyFromRecordset lrs
'                wsB.Cells(2, 14) = Now
'                Set lrs = Nothing
                
                Set lrs = New ADODB.Recordset
                lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
                lrs.Delete
                lrs.Update
                Set lrs = Nothing
                
                Set lrs = New ADODB.Recordset
                sql = " INSERT INTO Baixa "
                sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada, edicao, saida) "
                sql = sql & " VALUES "
                sql = sql & " ('" & ws.Cells(lin_inicio, 2) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 3) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 4) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 5) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 6) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 7) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 8) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 9) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 10) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 11) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 12) & "', "
                sql = sql & " '" & ws.Cells(lin_inicio, 13) & "', "
                sql = sql & " '" & Now & "') "
                
                lrs.Open sql, gConexao
                lrs.Close
                Set lrs = Nothing
                lsDesconectar
                
            Else
                MsgBox "VALOR INFORMADO MAIOR QUE ESTOQUE" & vbCrLf & vbCrLf & "VERIFIQUE A QUANTIDADE INFORMADA"
                GoTo ERRO
            End If
            result = MsgBox("BAIXA REALIZADA COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA BAIXAR OUTRO ITEM ?", vbYesNo + vbInformation)
            If result = vbYes Then
                GoTo NOVO
            Else
                TextBox201 = ""
                TextBox202 = ""
                TextBox203 = ""
                OptionButton1.Value = False
                OptionButton2.Value = False
                OptionButton3.Value = False
                ListBox1.Clear
                ListBox2.Clear
            End If
        End If
    End If
FIM:
    Application.ScreenUpdating = True
End Sub
