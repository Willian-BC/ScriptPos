'Option Explicit
Dim gConexao As New ADODB.Connection
Dim lrs As New ADODB.Recordset
Dim strConexao, sql As String
Dim ws, wsB, wsC As Worksheet
Dim wb As Workbook

Private Sub lsConectar()
    Set gConexao = New ADODB.Connection
    
    strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\Grupos\CZ1 - Transferencia Informacoes\10. Métodos e Processos\BD_Pin_BSA\Database_PIN_BSA.accdb;Persist Security Info=False"
    gConexao.Open strConexao
    
    If gConexao.State = adStateClosed Then
        MsgBox " Conexão falhou tente novamente", vbCritical
        Exit Sub
    End If
End Sub

Private Sub lsDesconectar()
    If gConexao.State = adStateOpen Then
        gConexao.Close
        Set gConexao = Nothing
    End If
End Sub

Private Sub UserForm_Initialize()
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f5 = fs.GetFile("C:\primus\SAX5.CSV")
    Set f6 = fs.GetFile("C:\primus\SAX6.CSV")
    
    Label101.Visible = True
    Label102.Visible = True
    Label103.Visible = False
    Label104.Visible = False
    TextBox100.Visible = False
    
    If f5.DateLastModified < f6.DateLastModified Then i = f5.DateLastModified Else i = f6.DateLastModified  'TextBox100 = f.DateCreated
    TextBox100 = i
    
    Sheets("Dados Primus").Visible = xlSheetVeryHidden
    Sheets("Componentes").Visible = xlSheetVeryHidden
    Sheets("Componentes (2)").Visible = xlSheetVeryHidden
    Sheets("Uniao").Visible = xlSheetVeryHidden
    Sheets("BD").Visible = xlSheetVeryHidden
    Sheets("Carreg").Visible = xlSheetVeryHidden
    Sheets("Desab").Visible = xlSheetVeryHidden
    If (Now - Sheets("Planilha1").Cells(3, 3)) < TimeValue("00:00:30") Then
        Sheets("Planilha1").Range("G2") = TimeValue("00:00:30") - (Now - Sheets("Planilha1").Cells(3, 3))
        Do While Sheets("Planilha1").Range("G2") <> 0
            Application.Wait (Now + TimeValue("00:00:01"))
            Sheets("Planilha1").Range("G2") = Sheets("Planilha1").Range("G2") - TimeValue("00:00:01")
        Loop
        On Error Resume Next
        Unload Me
    End If
End Sub

Private Sub CommandButton100_Click()
    Dim MyValue As Variant
    MyValue = InputBox("Digite a senha")
    If MyValue = "963" Then
        Sheets("Dados Primus").Visible = xlSheetVisible
        Sheets("Componentes").Visible = xlSheetVisible
        Sheets("Componentes (2)").Visible = xlSheetVisible
        Sheets("Uniao").Visible = xlSheetVisible
        Sheets("BD").Visible = xlSheetVisible
        Sheets("Carreg").Visible = xlSheetVisible
        Sheets("Desab").Visible = xlSheetVisible
    Else
        MsgBox ("Senha Incorreta")
    End If
End Sub

Private Sub CommandButton200_Click()
    ThisWorkbook.Save
End Sub

Private Sub CommandButton300_Click()
'ATUALIZAR
    ThisWorkbook.RefreshAll
    Sheets("Planilha1").Range("C3") = Now
    Unload Me
End Sub

Private Sub teste()
    If (Now - CDate(TextBox100)) > TimeValue("05:00:00") Or CDate(TextBox100) > Sheets("Planilha1").Cells(3, 3) Then
        Label101.Visible = False
        Label102.Visible = False
        Label103.Visible = True
        Label104.Visible = True
        TextBox100.Visible = True
    End If
End Sub

Private Sub CommandButton201_Click()
'CADASTRO OK
    If TextBox201 = "" Or TextBox202 = "" Or TextBox205 = "" Or TextBox207 = "" Then
        MsgBox "FAVOR PREENCHER AS INFORMAÇÕES!", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    If OptionButton1 = True Then
        lsConectar
        Set lrs = New ADODB.Recordset
        lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(TextBox200), gConexao, adOpenKeyset, adLockPessimistic
        
        lrs!Item = TextBox202
        lrs!descricao = TextBox203
        lrs!cor = TextBox204
        lrs!pedido = TextBox205
        lrs!ordem_prod = TextBox201
        lrs!cliente = TextBox206
        lrs!qt_pecas = TextBox207
        lrs!area_estoque = TextBox208
        lrs!comentario = TextBox209
        lrs!edicao = Now
        
        lrs.Update
        lrs.Close
        Set lrs = Nothing
        lsDesconectar
        
        MsgBox "CADASTRO ATUALIZADO COM SUCESSO!", vbInformation
        ThisWorkbook.Save
        CommandButton202_Click
        
        UserForm1.MultiPage1.Value = 1
    Else
        lsConectar
        Set lrs = New ADODB.Recordset
        
        sql = " INSERT INTO BD_dados "
        sql = sql & " (item, descricao, cor, pedido, ordem_prod, cliente, qt_pecas, area_estoque, comentario, entrada) "
        sql = sql & " VALUES "
        sql = sql & " ('" & TextBox202 & "', "
        sql = sql & " '" & TextBox203 & "', "
        sql = sql & " '" & TextBox204 & "', "
        sql = sql & " '" & TextBox205 & "', "
        sql = sql & " '" & TextBox201 & "', "
        sql = sql & " '" & TextBox206 & "', "
        sql = sql & " '" & TextBox207 & "', "
        sql = sql & " '" & TextBox208 & "', "
        sql = sql & " '" & TextBox209 & "', "
        sql = sql & " '" & Now & "') "
        
        lrs.Open sql, gConexao
        Set lrs = Nothing
        lsDesconectar
        
        MsgBox "CADASTRO REALIZADO COM SUCESSO!", vbInformation
        ThisWorkbook.Save
        CommandButton202_Click
        
        TextBox201.SetFocus
    End If
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub TextBox201_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'CADASTRO PESQUISAR OP
    If TextBox201 = "" Then
        CommandButton202_Click
    Else
        Set wsC = Sheets("Componentes")
        On Error Resume Next
        wsC.Cells(2, 11) = CLng(TextBox201)
        TextBox202 = wsC.Cells(3, 11)
        TextBox203 = wsC.Cells(4, 11)
        TextBox204 = wsC.Cells(5, 11)
        TextBox205 = wsC.Cells(6, 11)
        TextBox206 = wsC.Cells(7, 11)
    End If
    teste
End Sub

Private Sub CommandButton202_Click()
'CADASTRO LIMPAR
    TextBox201 = ""
    TextBox202 = ""
    TextBox203 = ""
    TextBox204 = ""
    TextBox205 = ""
    TextBox206 = ""
    TextBox207 = ""
    TextBox208 = ""
    TextBox209 = ""
End Sub

Private Sub TextBox301_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    a = TextBox301
    CommandButton302_Click
    TextBox301 = a
End Sub

Private Sub TextBox305_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'SALDO PESQUISAR OP
    If TextBox305 = "" Then
        CommandButton302_Click
    ElseIf TextBox301 = "" Then
        GoTo FIM
    Else
        Set ws = Sheets("Componentes")
        Set wsC = Sheets("Componentes (2)")
        ws.Cells(2, 11) = CLng(TextBox301)
        
        If ws.Cells(5, 11) = "" Then
            On Error Resume Next
            wsC.ListObjects(1).ShowAutoFilter = True
            wsC.ListObjects(1).AutoFilter.ShowAllData
            
            wsC.ListObjects(1).Range.AutoFilter Field:=6, Criteria1:=TextBox305
            wsC.ListObjects(1).Range.AutoFilter Field:=4, Criteria1:=ws.Cells(7, 11)
            wsC.ListObjects(1).Range.AutoFilter Field:=3, Criteria1:="=*" & Split(ws.Cells(4, 11), "   ")(0) & "*"
            TextBox304 = Right(wsC.Cells(wsC.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row, 3), 3)
        Else
            TextBox304 = ws.Cells(5, 11)
        End If
        
        TextBox302 = ws.Cells(3, 11)
        TextBox303 = ws.Cells(4, 11)
        TextBox306 = ws.Cells(7, 11)
        TextBox307 = ws.Cells(8, 11)
        TextBox308 = ws.Cells(9, 11)
        
        If TextBox302 <> "NÃO ENCONTRADO" Then
            lsConectar
            Set lrs = New ADODB.Recordset
            
            sql = " INSERT INTO Consulta "
            sql = sql & " (item, descricao, cor, pedido, ordem_prod, programa, cliente, saldo, horario) "
            sql = sql & " VALUES "
            sql = sql & " ('" & TextBox302 & "', "
            sql = sql & " '" & TextBox303 & "', "
            sql = sql & " '" & TextBox304 & "', "
            sql = sql & " '" & TextBox306 & "', "
            sql = sql & " '" & TextBox301 & "', "
            sql = sql & " '" & TextBox305 & "', "
            sql = sql & " '" & TextBox307 & "', "
            sql = sql & " '" & TextBox308 & "', "
            sql = sql & " '" & Now & "') "
            
            lrs.Open sql, gConexao
            Set lrs = Nothing
            lsDesconectar
        End If
    End If
    teste
FIM:
End Sub

Private Sub CommandButton302_Click()
'SALDO LIMPAR
    TextBox301 = ""
    TextBox302 = ""
    TextBox303 = ""
    TextBox304 = ""
    TextBox305 = ""
    TextBox306 = ""
    TextBox307 = ""
    TextBox308 = ""
End Sub

Private Sub TextBox401_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If OptionButton2 = False Then
        a = TextBox401
        CommandButton402_Click
        TextBox401 = a
    End If
End Sub

Private Sub TextBox405_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'CARREG - IMPORTAÇÃO DADOS ATRAVES O.P.
    If TextBox405 = "" Then
        CommandButton402_Click
    ElseIf TextBox401 = "" Then
        GoTo FIM
    Else
        Set ws = Sheets("Componentes")
        Set wsC = Sheets("Componentes (2)")
        ws.Cells(2, 11) = CLng(TextBox401)
        
        If ws.Cells(5, 11) = "" Then
            On Error Resume Next
            wsC.ListObjects(1).ShowAutoFilter = True
            wsC.ListObjects(1).AutoFilter.ShowAllData
            
            wsC.ListObjects(1).Range.AutoFilter Field:=6, Criteria1:=TextBox405
            wsC.ListObjects(1).Range.AutoFilter Field:=4, Criteria1:=ws.Cells(7, 11)
            wsC.ListObjects(1).Range.AutoFilter Field:=3, Criteria1:="=*" & Split(ws.Cells(4, 11), "   ")(0) & "*"
            TextBox404 = Right(wsC.Cells(wsC.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row, 3), 3)
        Else
            TextBox404 = ws.Cells(5, 11)
        End If
        
        TextBox402 = ws.Cells(3, 11)
        TextBox403 = ws.Cells(4, 11)
        TextBox406 = ws.Cells(7, 11)
        TextBox407 = ws.Cells(8, 11)
    End If
FIM:
End Sub

Private Sub CommandButton401_Click()
'CARREG - OK
    If TextBox401 = "" Or TextBox402 = "" Or TextBox406 = "" Or TextBox408 = "" Or TextBox409 = "" Or TextBox410 = "" Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS PARA CADASTRO!", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    If OptionButton2 = True Then
        lsConectar
        Set lrs = New ADODB.Recordset
        lrs.Open "SELECT * FROM Carreg WHERE ID = " & CInt(TextBox400), gConexao, adOpenKeyset, adLockPessimistic
        
        lrs!Item = TextBox402
        lrs!descricao = TextBox403
        lrs!cor = TextBox404
        lrs!pedido = TextBox406
        lrs!ordem_prod = TextBox401
        lrs!programa = TextBox405
        lrs!cliente = TextBox407
        lrs!qt_pecas = TextBox408
        lrs!peso = TextBox409
        lrs!turno = TextBox410
        lrs!comentario = TextBox411
        lrs!edicao = Now
        
        lrs.Update
        lrs.Close
        Set lrs = Nothing
        lsDesconectar
        
        MsgBox "CADASTRO ATUALIZADO COM SUCESSO!", vbInformation
        CommandButton402_Click
        UserForm1.MultiPage1.Value = 6
    Else
NOVO:
        lsConectar
        Set lrs = New ADODB.Recordset
        
        sql = " INSERT INTO Carreg "
        sql = sql & " (item, descricao, cor, pedido, ordem_prod, programa, cliente, qt_pecas, peso, turno, comentario, entrada) "
        sql = sql & " VALUES "
        sql = sql & " ('" & TextBox402 & "', "
        sql = sql & " '" & TextBox403 & "', "
        sql = sql & " '" & TextBox404 & "', "
        sql = sql & " '" & TextBox406 & "', "
        sql = sql & " '" & TextBox401 & "', "
        sql = sql & " '" & TextBox405 & "', "
        sql = sql & " '" & TextBox407 & "', "
        sql = sql & " '" & TextBox408 & "', "
        sql = sql & " '" & TextBox409 & "', "
        sql = sql & " '" & TextBox410 & "', "
        sql = sql & " '" & TextBox411 & "', "
        If TimeValue(Now) > TimeValue("00:00:00") And TimeValue(Now) < TimeValue("06:00:00") Then
            sql = sql & " '" & Now - 1 & "') "
        Else
            sql = sql & " '" & Now & "') "
        End If
        
        lrs.Open sql, gConexao
        Set lrs = Nothing
        lsDesconectar
        
        If TextBox412 <> "" Then
            TextBox412 = TextBox412 - 1
            If TextBox412 > 0 Then GoTo NOVO
        End If
        
        MsgBox "CADASTRO REALIZADO COM SUCESSO!", vbOKOnly + vbInformation
        CheckBox2 = False
        CommandButton402_Click
        TextBox401.SetFocus
    End If
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub CommandButton402_Click()
'CARREG - LIMPAR
    TextBox400 = ""
    TextBox401 = ""
    TextBox402 = ""
    TextBox403 = ""
    TextBox404 = ""
    TextBox405 = ""
    TextBox406 = ""
    TextBox407 = ""
    TextBox408 = ""
    TextBox409 = ""
    TextBox410 = ""
    TextBox411 = ""
    OptionButton2 = False
End Sub

Private Sub CheckBox2_Click()
'CARREG
    If CheckBox2 = True Then
        Label412.Visible = True
        TextBox412.Visible = True
    Else
        Label412.Visible = False
        TextBox412.Visible = False
        TextBox412 = ""
    End If
End Sub

Private Sub TextBox501_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'DESAB - IMPORTAÇÃO DADOS ATRAVES O.P.
    If TextBox501 = "" Then
        CommandButton502_Click
    Else
        Set wsC = Sheets("Componentes (2)")
        
        On Error Resume Next
        wsC.ListObjects(1).ShowAutoFilter = True
        wsC.ListObjects(1).AutoFilter.ShowAllData
        
        wsC.ListObjects(1).Range.AutoFilter Field:=5, Criteria1:=TextBox501
        lin = wsC.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        TextBox502 = wsC.Cells(lin, 2)
        TextBox503 = wsC.Cells(lin, 3)
        TextBox504 = Right(wsC.Cells(wsC.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row, 3), 2)
        TextBox505 = wsC.Cells(lin, 6)
        TextBox506 = wsC.Cells(lin, 4)
        TextBox507 = wsC.Cells(lin, 7)
    End If
End Sub

Private Sub CommandButton501_Click()
'DESAB - OK
    If TextBox501 = "" Or TextBox502 = "" Or TextBox506 = "" Or TextBox508 = "" Or TextBox509 = "" Or TextBox510 = "" Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS PARA CADASTRO!", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    If OptionButton3 = True Then
        lsConectar
        Set lrs = New ADODB.Recordset
        lrs.Open "SELECT * FROM Desab WHERE ID = " & CInt(TextBox500), gConexao, adOpenKeyset, adLockPessimistic
        
        lrs!Item = TextBox502
        lrs!descricao = TextBox503
        lrs!cor = TextBox504
        lrs!pedido = TextBox506
        lrs!ordem_prod = TextBox501
        lrs!programa = TextBox505
        lrs!cliente = TextBox507
        lrs!qt_pecas = TextBox508
        lrs!peso = TextBox509
        lrs!turno = TextBox510
        lrs!comentario = TextBox511
        lrs!edicao = Now
        
        lrs.Update
        lrs.Close
        Set lrs = Nothing
        lsDesconectar
        
        MsgBox "CADASTRO ATUALIZADO COM SUCESSO!", vbInformation
        CommandButton502_Click
        UserForm1.MultiPage1.Value = 6
    Else
NOVO:
        lsConectar
        Set lrs = New ADODB.Recordset
        
        sql = " INSERT INTO Desab "
        sql = sql & " (item, descricao, cor, pedido, ordem_prod, programa, cliente, qt_pecas, peso, turno, comentario, saida) "
        sql = sql & " VALUES "
        sql = sql & " ('" & TextBox502 & "', "
        sql = sql & " '" & TextBox503 & "', "
        sql = sql & " '" & TextBox504 & "', "
        sql = sql & " '" & TextBox506 & "', "
        sql = sql & " '" & TextBox501 & "', "
        sql = sql & " '" & TextBox505 & "', "
        sql = sql & " '" & TextBox507 & "', "
        sql = sql & " '" & TextBox508 & "', "
        sql = sql & " '" & TextBox509 & "', "
        sql = sql & " '" & TextBox510 & "', "
        sql = sql & " '" & TextBox511 & "', "
        If TimeValue(Now) > TimeValue("00:00:00") And TimeValue(Now) < TimeValue("06:00:00") Then
            sql = sql & " '" & Now - 1 & "') "
        Else
            sql = sql & " '" & Now & "') "
        End If
        
        lrs.Open sql, gConexao
        Set lrs = Nothing
        lsDesconectar
        
        If TextBox512 <> "" Then
            TextBox512 = TextBox512 - 1
            If TextBox512 > 0 Then GoTo NOVO
        End If
        
        MsgBox "CADASTRO REALIZADO COM SUCESSO!", vbOKOnly + vbInformation
        CheckBox3 = False
        CommandButton502_Click
        TextBox501.SetFocus
    End If
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub CommandButton502_Click()
'DESAB - LIMPAR
    TextBox500 = ""
    TextBox501 = ""
    TextBox502 = ""
    TextBox503 = ""
    TextBox504 = ""
    TextBox505 = ""
    TextBox506 = ""
    TextBox507 = ""
    TextBox508 = ""
    TextBox509 = ""
    TextBox510 = ""
    TextBox511 = ""
    OptionButton3 = False
End Sub

Private Sub CheckBox3_Click()
'DESAB
    If CheckBox3 = True Then
        Label512.Visible = True
        TextBox512.Visible = True
    Else
        Label512.Visible = False
        TextBox512.Visible = False
        TextBox512 = ""
    End If
End Sub

Private Sub CheckBox1_Click()
    If CheckBox1 = True Then
        If TimeValue(Now) > TimeValue("00:00:00") And TimeValue(Now) < TimeValue("06:00:00") Then
            TextBox601 = Date - 1
        Else
            TextBox601 = Date
        End If
    Else
        TextBox601 = ""
    End If
End Sub

Private Sub CommandButton601_Click()
'RELATÓRIO - OK
    If TextBox601 = "" Or TextBox602 = "" Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS PARA CONTINUAR!", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Set ws = Sheets("Carreg")
    Set wsB = Sheets("Desab")
    
    If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    ws.ShowAllData
    ws.Rows("2:" & ws.Cells(1, 1).End(xlDown).Row).ClearContents
    
    If Not wsB.AutoFilterMode Then wsB.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    wsB.ShowAllData
    wsB.Rows("2:" & wsB.Cells(1, 1).End(xlDown).Row).ClearContents
    
    lsConectar
    Set lrs = New ADODB.Recordset
    lrs.Open " SELECT * FROM Carreg ", gConexao, adOpenKeyset, adLockPessimistic
    ws.Cells(2, 1).CopyFromRecordset lrs
    ws.Columns("A:A").NumberFormat = "0"
    ws.Columns("I:J").NumberFormat = "0"
    ws.Columns("M:M").NumberFormat = "dd/mm/yyyy"
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
    
    lsConectar
    Set lrs = New ADODB.Recordset
    lrs.Open " SELECT * FROM Desab ", gConexao, adOpenKeyset, adLockPessimistic
    wsB.Cells(2, 1).CopyFromRecordset lrs
    wsB.Columns("A:A").NumberFormat = "0"
    wsB.Columns("I:J").NumberFormat = "0"
    wsB.Columns("M:M").NumberFormat = "dd/mm/yyyy"
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
    
    Set rngAF = ws.Range("A1:A" & ws.Cells(1, 1).End(xlDown).Row)
    Set rngAFB = wsB.Range("A1:A" & wsB.Cells(1, 1).End(xlDown).Row)
    
    ws.Range("M:M").AutoFilter Field:=13, Criteria1:=Format(CDate(TextBox601), "dd/mm/yyyy")
    ws.Range("K:K").AutoFilter Field:=11, Criteria1:=TextBox602
    wsB.Range("M:M").AutoFilter Field:=13, Criteria1:=Format(CDate(TextBox601), "dd/mm/yyyy")
    wsB.Range("K:K").AutoFilter Field:=11, Criteria1:=TextBox602
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_inicioB = wsB.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    If ws.Cells(lin_inicio, 1).Value = 0 And wsB.Cells(lin_inicioB, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        ListBox2.Clear
        ListBox3.Clear
        ws.ShowAllData
        wsB.ShowAllData
        GoTo FIM
    End If
    
    If ws.Cells(lin_inicio, 1).Value <> 0 Then
        Dim arrayItems2()
        With Planilha4
            ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, ws.Range("A:A")), 1 To 13)
            Me.ListBox2.ColumnCount = 13
            Me.ListBox2.ColumnWidths = "0;65;140;30;0;60;40;100;40;40;40;60;85"
            i = 0
            For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
                Me.ListBox2.AddItem
                For coluna = 1 To 13
                    arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
                Next coluna
                i = i + 1
            Next rngcell
            Me.ListBox2.List = arrayItems2()
        End With
    End If
    If wsB.Cells(lin_inicioB, 1).Value <> 0 Then
        Dim arrayItems3()
        With Planilha7
            ReDim arrayItems3(0 To WorksheetFunction.Subtotal(102, wsB.Range("A:A")), 1 To 13)
            Me.ListBox3.ColumnCount = 13
            Me.ListBox3.ColumnWidths = "0;65;140;30;0;60;40;100;40;40;40;60;85"
            i = 0
            For Each rngcell In rngAFB.SpecialCells(xlCellTypeVisible)
                Me.ListBox3.AddItem
                For coluna = 1 To 13
                    arrayItems3(i, coluna) = .Cells(rngcell.Row, coluna).Value
                Next coluna
                i = i + 1
            Next rngcell
            Me.ListBox3.List = arrayItems3()
        End With
    End If
    
    TextBox603 = Format(WorksheetFunction.Subtotal(109, ws.Range("I:I")), "0")
    TextBox604 = Format(WorksheetFunction.Subtotal(109, ws.Range("J:J")), "#,##0.00")
    TextBox605 = Format(WorksheetFunction.Subtotal(109, wsB.Range("I:I")), "0")
    TextBox606 = Format(WorksheetFunction.Subtotal(109, wsB.Range("J:J")), "#,##0.00")
FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton602_Click()
'RELATÓRIO - LIMPAR
    TextBox601 = ""
    TextBox602 = ""
    TextBox603 = ""
    TextBox604 = ""
    TextBox605 = ""
    TextBox606 = ""
    CheckBox1 = False
    ListBox2.Clear
    ListBox3.Clear
End Sub

Private Sub CommandButton605_Click()
'RELATÓRIO - CARREG EDITAR
    If ListBox2.ListCount = 0 Then
        MsgBox "FAVOR REALIZAR UMA PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    On Error Resume Next
    
    Set ws = Sheets("Carreg")
    ID = ListBox2.List(ListBox2.ListIndex, 0)
    
    If ID = 0 Or Not IsNumeric(ID) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA EDITAR", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    On Error Resume Next
    lin_inicio = Application.WorksheetFunction.Match(CInt(ID), ws.Range("A:A"), 0)
    
    TextBox400 = ws.Cells(lin_inicio, 1)
    TextBox401 = ws.Cells(lin_inicio, 6)
    TextBox402 = ws.Cells(lin_inicio, 2)
    TextBox403 = ws.Cells(lin_inicio, 3)
    TextBox404 = ws.Cells(lin_inicio, 4)
    TextBox405 = ws.Cells(lin_inicio, 7)
    TextBox406 = ws.Cells(lin_inicio, 5)
    TextBox407 = ws.Cells(lin_inicio, 8)
    TextBox408 = ws.Cells(lin_inicio, 9)
    TextBox409 = ws.Cells(lin_inicio, 10)
    TextBox410 = ws.Cells(lin_inicio, 11)
    TextBox411 = ws.Cells(lin_inicio, 12)
    OptionButton2 = True
    
    CommandButton602_Click
    UserForm1.MultiPage1.Value = 3
FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton606_Click()
'RELATÓRIO - CARREG DELETAR
    If ListBox2.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    On Error Resume Next
    ID = ListBox2.List(ListBox2.ListIndex, 0)
    If ID = 0 Or Not IsNumeric(ID) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA DELETAR", vbCritical
        GoTo FIM
    End If
    
    result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR O ITEM ?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
    If result = vbYes Then
        
        lsConectar
        Set lrs = New ADODB.Recordset
        lrs.Open "SELECT * FROM Carreg WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
        lrs.Delete
        lrs.Update
        Set lrs = Nothing
        
        MsgBox "CADASTRO EXCLUIDO COM SUCESSO!", vbInformation
        ListBox2.Clear
    End If
FIM:
End Sub

Private Sub CommandButton603_Click()
'RELATÓRIO - DESAB EDITAR
    If ListBox3.ListCount = 0 Then
        MsgBox "FAVOR REALIZAR UMA PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    On Error Resume Next
    
    Set ws = Sheets("Desab")
    ID = ListBox3.List(ListBox3.ListIndex, 0)
    
    If ID = 0 Or Not IsNumeric(ID) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA EDITAR", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    On Error Resume Next
    lin_inicio = Application.WorksheetFunction.Match(CInt(ID), ws.Range("A:A"), 0)
    
    TextBox500 = ws.Cells(lin_inicio, 1)
    TextBox501 = ws.Cells(lin_inicio, 6)
    TextBox502 = ws.Cells(lin_inicio, 2)
    TextBox503 = ws.Cells(lin_inicio, 3)
    TextBox504 = ws.Cells(lin_inicio, 4)
    TextBox505 = ws.Cells(lin_inicio, 7)
    TextBox506 = ws.Cells(lin_inicio, 5)
    TextBox507 = ws.Cells(lin_inicio, 8)
    TextBox508 = ws.Cells(lin_inicio, 9)
    TextBox509 = ws.Cells(lin_inicio, 10)
    TextBox510 = ws.Cells(lin_inicio, 11)
    TextBox511 = ws.Cells(lin_inicio, 12)
    OptionButton3 = True
    
    CommandButton602_Click
    UserForm1.MultiPage1.Value = 5
FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton604_Click()
'RELATÓRIO - DESAB DELETAR
    If ListBox3.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    On Error Resume Next
    ID = ListBox3.List(ListBox3.ListIndex, 0)
    If ID = 0 Or Not IsNumeric(ID) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA DELETAR", vbCritical
        GoTo FIM
    End If
    
    result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR O ITEM ?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
    If result = vbYes Then
        
        lsConectar
        Set lrs = New ADODB.Recordset
        lrs.Open "SELECT * FROM Desab WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
        lrs.Delete
        lrs.Update
        Set lrs = Nothing
        
        MsgBox "CADASTRO EXCLUIDO COM SUCESSO!", vbInformation
        ListBox3.Clear
    End If
FIM:
End Sub

Private Sub CommandButton701_Click()
'OK APONTAMENTO - EXIBIR DADOS EM TELA

    If TextBox701 = "" Or TextBox702 = "" Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS!", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Set wsC = Sheets("Componentes (2)")
    
    On Error Resume Next
    wsC.ListObjects(1).ShowAutoFilter = True
    wsC.ListObjects(1).AutoFilter.ShowAllData
    Set rngAF = wsC.Range("A1:A" & wsC.Cells(1, 1).End(xlDown).Row)
    
    wsC.ListObjects(1).Range.AutoFilter Field:=4, Criteria1:=TextBox701
    wsC.ListObjects(1).Range.AutoFilter Field:=3, Criteria1:="=*" & TextBox702 & "*"
    
    lin_inicio = wsC.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = wsC.Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If wsC.Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        wsC.ListObjects(1).AutoFilter.ShowAllData
        GoTo FIM
    End If
    
    Dim arrayItems2()
    With wsC
        ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, wsC.Range("A:A")), 1 To 10)
        Me.ListBox4.ColumnCount = 10
        Me.ListBox4.ColumnWidths = ";130;350;;;;200;;;"
        i = 0
        For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
            Me.ListBox3.AddItem
            For coluna = 1 To 10
                arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
            Next coluna
            i = i + 1
        Next rngcell
        Me.ListBox4.List = arrayItems2()
    End With
FIM:
End Sub

Private Sub CommandButton702_Click()
'LIMPAR APONTAMENTO

    TextBox701 = ""
    TextBox702 = ""
    TextBox703 = ""
    TextBox704 = ""
    ListBox4.Clear
End Sub

Private Sub TextBox703_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'APONTAMENTO - COLETAR O.P. ATRAVÉS DO INDICE

    If TextBox703 = "" Or ListBox4.ListCount = 0 Then
        TextBox704 = ""
    Else
        Application.ScreenUpdating = False
        Set wsC = Sheets("Componentes (2)")
        
        wsC.ListObjects(1).Range.AutoFilter Field:=1, Criteria1:=TextBox703
        lin_inicio = wsC.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        If wsC.Cells(lin_inicio, 1).Value = 0 Then
            wsC.ListObjects(1).Range.AutoFilter Field:=1
            TextBox704 = ""
            GoTo FIM
        End If
        TextBox704 = wsC.Cells(lin_inicio, 5)
        TextBox704.SetFocus
    End If
FIM:
End Sub

Private Sub TextBox601_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextBox601.MaxLength = 10
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        If TextBox601.SelStart = 2 Then
            TextBox601.SelText = "/"
        End If
        If TextBox601.SelStart = 5 Then
            TextBox601.SelText = "/"
        End If
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox305_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TextBox405_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TextBox409_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 44 'Aceita VIRGULA
        Case 48 To 57
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub
Private Sub TextBox410_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TextBox411_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TextBox412_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub
Private Sub TextBox509_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 44 'Aceita VIRGULA
        Case 48 To 57
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub
Private Sub TextBox510_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TextBox511_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TextBox512_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub
Private Sub TextBox602_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub TextBox702_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
