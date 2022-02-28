Private Sub CommandButton5_Click()
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Set ws = Sheets("Base de dados")
    Set wsB = Sheets("Baixados")
    
    result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR UM REGISTRO?", vbYesNo + vbCritical)
    If result = vbYes Then
ERRO:
        ID = Application.InputBox("INFORME O ID")
        If ID = 0 Then GoTo FIM
        
        result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR O ID " & ID & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
        If result = vbYes Then
        
            Application.ScreenUpdating = False
            If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
            On Error Resume Next
            ws.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
            
            lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            
            If ws.Cells(lin_inicio, 1).Value = 0 Then
                MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR O NÚMERO DE ID", vbCritical
                ws.Range("A:A").AutoFilter Field:=1
                GoTo ERRO
            End If
            
            lsConectar
            Set lrs = New ADODB.Recordset
            lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
            lrs.Delete
            lrs.Update
            Set lrs = Nothing
            
            Set lrs = New ADODB.Recordset
            sql = " INSERT INTO Excluir "
            sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada, edicao, excluido) "
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
            
            MsgBox "CADASTRO EXCLUIDO COM SUCESSO!", vbInformation
            ListBox1.Clear
        End If
    End If
FIM:
    Application.ScreenUpdating = True
End Sub
