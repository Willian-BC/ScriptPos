'Microsoft ActiveX Data Objects 6.1 Library
'Option Explicit
Dim gConexao As New ADODB.Connection
Dim lrs As New ADODB.Recordset
Dim strConexao, sql As String
Dim ws, wsB, wsC As Worksheet
Dim wb As Workbook

Private Sub lsConectar()
    Set gConexao = New ADODB.Connection
    
    strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\Grupos\CZ1 - Transferencia Informacoes\10. Métodos e Processos\BD_Almox\Database_ALM.accdb;Persist Security Info=False"
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
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    Sheets("Baixa").Visible = xlSheetVeryHidden
    Sheets("Excluido").Visible = xlSheetVeryHidden
    Sheets("Descrição").Visible = xlSheetVeryHidden
End Sub

Private Sub CommandButton100_Click()
    Dim MyValue As Variant
    MyValue = InputBox("Digite a senha")
    If MyValue = "1010" Then
        Sheets("Base de dados").Visible = xlSheetVisible
        Sheets("Baixa").Visible = xlSheetVisible
        Sheets("Excluido").Visible = xlSheetVisible
        Sheets("Descrição").Visible = xlSheetVisible
    Else
        MsgBox ("Senha Incorreta")
    End If
End Sub

Private Sub CommandButton200_Click()
    ThisWorkbook.Save
End Sub

Private Sub OptionButton3_Click()
    ListBox1.Clear
    TextBox4 = ""
End Sub
Private Sub OptionButton4_Click()
    ListBox1.Clear
    TextBox4 = ""
End Sub
Private Sub OptionButton5_Click()
    ListBox1.Clear
    TextBox4 = ""
End Sub

Private Sub CommandButton1_Click()
'OK PESQUISA'
    
    If OptionButton3 = True Then
        Set ws = Sheets("Base de dados")
    ElseIf OptionButton4 = True Then
        Set ws = Sheets("Baixa")
    Else
        Set ws = Sheets("Excluido")
    End If
    
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    
    If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    ws.ShowAllData
    ws.Rows("2:" & ws.Cells(1, 1).End(xlDown).Row).ClearContents
    
    
    lsConectar
    Set lrs = New ADODB.Recordset
    lrs.Open " SELECT * FROM BD_dados ", gConexao, adOpenKeyset, adLockPessimistic
    ws.Cells(2, 1).CopyFromRecordset lrs
    ws.Columns("A:A,D:D").NumberFormat = "0"
    ws.Columns("G:H").NumberFormat = "dd/mm/yyyy"
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
    
    Set rngAF = ws.Range("A1:A" & ws.Cells(1, 1).End(xlDown).Row)
    
    If TextBox1 <> "" Then ws.Range("B:B").AutoFilter Field:=2, Criteria1:=TextBox1.Text
    If TextBox2 <> "" Then ws.Range("C:C").AutoFilter Field:=3, Criteria1:="=*" & TextBox2.Text & "*"
    If TextBox3 <> "" Then ws.Range("E:E").AutoFilter Field:=5, Criteria1:=TextBox3.Text
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = ws.Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If ws.Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        ws.ShowAllData
        GoTo FIM
    Else
    
    Dim arrayItems2()
        With ws
            ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, ws.Range("A:A")), 1 To 7)
            Me.ListBox1.ColumnCount = 7
            Me.ListBox1.ColumnWidths = "40;100;300;70;70;200;70"
            i = 0
            For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
                Me.ListBox1.AddItem
                For coluna = 1 To 7
                    arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
                Next coluna
                i = i + 1
            Next rngcell
            Me.ListBox1.List = arrayItems2()
        End With
    End If
    
    TextBox4 = Format(WorksheetFunction.Subtotal(109, ws.Range("D:D")), "#,##0")
    
    If CheckBox1 = True Then
        Set rngAJ = ws.Range("B1:H" & lin_fim).SpecialCells(xlCellTypeVisible)
        rngAJ.Copy
        Workbooks.Add
        Range("A1").PasteSpecial Paste:=xlPasteValues
    End If
    
FIM:
    Application.ScreenUpdating = True
    
    If CheckBox1 = True Then
        result = MsgBox("DADOS EXPORTADOS COM SUCESSO !" & vbCrLf & "DESEJA FECHAR O FORMULÁRIO ?" & vbCrLf & vbCrLf & "É NECESSÁRIO FECHAR PARA EDITAR OS DADOS", vbYesNo + vbInformation)
        If result = vbYes Then
            Unload Me
        Else
            ActiveWindow.WindowState = xlMinimized
            wb.Activate
            CheckBox1 = False
        End If
    End If
    
End Sub

Private Sub CommandButton2_Click()
'LIMPAR PESQUISA
    TextBox1 = ""
    TextBox2 = ""
    TextBox3 = ""
    TextBox4 = ""
    ListBox1.Clear
    CheckBox1 = False
    OptionButton3 = True
    OptionButton4 = False
    OptionButton5 = False
End Sub

Private Sub CommandButton4_Click()
'EDITAR PESQUISA
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR REALIZAR UMA PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    If OptionButton4 = True Then
        MsgBox "NÃO É POSSÍVEL EDITAR UM ITEM BAIXADO !" & vbCrLf & vbCrLf & "FAVOR ALTERAR O TIPO DE PESQUISA", vbCritical
        GoTo FIM
    ElseIf OptionButton5 = True Then
        MsgBox "NÃO É POSSÍVEL EDITAR UM ITEM EXCLUIDO !" & vbCrLf & vbCrLf & "FAVOR ALTERAR O TIPO DE PESQUISA", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Set ws = Sheets("Base de dados")
    ID = ListBox1.List(ListBox1.ListIndex, 0)
    If ID = 0 Or Not IsNumeric(ID) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA EDITAR", vbCritical
        GoTo FIM
    End If
    
    If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    ws.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    TextBox101 = ws.Cells(lin_inicio, 1)
    TextBox102 = ws.Cells(lin_inicio, 2)
    TextBox103 = ws.Cells(lin_inicio, 3)
    TextBox104 = ws.Cells(lin_inicio, 4)
    TextBox105 = ws.Cells(lin_inicio, 5)
    TextBox106 = ws.Cells(lin_inicio, 6)
    
    ws.ShowAllData
    ListBox1.Clear
    UserForm1.MultiPage1.Value = 1
FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton5_Click()
'APAGAR PESQUISA
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR REALIZAR UMA PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    If OptionButton4 = True Then
        MsgBox "NÃO É POSSÍVEL DELETAR UM ITEM BAIXADO !" & vbCrLf & vbCrLf & "FAVOR ALTERAR O TIPO DE PESQUISA", vbCritical
        GoTo FIM
    ElseIf OptionButton5 = True Then
        MsgBox "NÃO É POSSÍVEL DELETAR UM ITEM EXCLUIDO !" & vbCrLf & vbCrLf & "FAVOR ALTERAR O TIPO DE PESQUISA", vbCritical
        GoTo FIM
    End If
    
    Set ws = Sheets("Base de dados")
    ID = ListBox1.List(ListBox1.ListIndex, 0)
    If ID = 0 Or Not IsNumeric(ID) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA DELETAR", vbCritical
        GoTo FIM
    End If
    
    result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR UM REGISTRO?", vbYesNo + vbCritical)
    If result = vbYes Then
        Application.ScreenUpdating = False
        If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
        On Error Resume Next
        ws.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
        lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        lsConectar
        Set lrs = New ADODB.Recordset
        lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
        lrs.Delete
        lrs.Update
        Set lrs = Nothing
        
        Set lrs = New ADODB.Recordset
        sql = " INSERT INTO Excluir "
        sql = sql & " (item, descricao, qt_pecas, area_estoque, comentario, entrada, edicao, excluido) "
        sql = sql & " VALUES "
        sql = sql & " ('" & ws.Cells(lin_inicio, 2) & "', "
        sql = sql & " '" & ws.Cells(lin_inicio, 3) & "', "
        sql = sql & " '" & ws.Cells(lin_inicio, 4) & "', "
        sql = sql & " '" & ws.Cells(lin_inicio, 5) & "', "
        sql = sql & " '" & ws.Cells(lin_inicio, 6) & "', "
        sql = sql & " '" & ws.Cells(lin_inicio, 7) & "', "
        sql = sql & " '" & ws.Cells(lin_inicio, 8) & "', "
        sql = sql & " '" & Now & "') "
        
        lrs.Open sql, gConexao
        lrs.Close
        Set lrs = Nothing
        lsDesconectar
        
        MsgBox "CADASTRO EXCLUIDO COM SUCESSO!", vbInformation
        ListBox1.Clear
        Application.ScreenUpdating = True
    End If
FIM:
End Sub

Private Sub CommandButton101_Click()
'OK CADASTRO'
    If TextBox101 <> "" Then
        MsgBox "CADASTRO JÁ EXISTE!" & vbCrLf & vbCrLf & "CLIQUE EM ATUALIZAR REGISTRO OU LIMPAR", vbCritical
        GoTo FIM
    ElseIf TextBox102 = "" Or TextBox105 = "" Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS PARA CADASTRO!", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    
    lsConectar
    Set lrs = New ADODB.Recordset
    
    sql = " INSERT INTO BD_dados "
    sql = sql & " (item, descricao, qt_pecas, area_estoque, comentario, entrada) "
    sql = sql & " VALUES "
    sql = sql & " ('" & TextBox102 & "', "
    sql = sql & " '" & TextBox103 & "', "
    sql = sql & " '" & TextBox104 & "', "
    sql = sql & " '" & TextBox105 & "', "
    sql = sql & " '" & TextBox106 & "', "
    sql = sql & " '" & Now & "') "
    
    lrs.Open sql, gConexao
    Set lrs = Nothing
    lsDesconectar
    
    result = MsgBox("CADASTRO REALIZADO COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA LIMPAR OS DADOS?", vbYesNo + vbInformation)
    If result = vbYes Then
        CommandButton102_Click
    End If
    
    TextBox102.SetFocus
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub CommandButton102_Click()
'LIMPAR CADASTRO
    TextBox101 = ""
    TextBox102 = ""
    TextBox103 = ""
    TextBox104 = ""
    TextBox105 = ""
    TextBox106 = ""
End Sub

Private Sub CommandButton104_Click()
'ATUALIZAR CADASTRO
    If TextBox101 = "" Then
        MsgBox "CADASTRO NÃO EXISTE!" & vbCrLf & vbCrLf & "CLIQUE EM CADASTRAR", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    
    lsConectar
    Set lrs = New ADODB.Recordset
    lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(TextBox101), gConexao, adOpenKeyset, adLockPessimistic
    
    lrs!Item = TextBox102
    lrs!descricao = TextBox103
    lrs!qt_pecas = TextBox104
    lrs!area_estoque = TextBox105
    lrs!comentario = TextBox106
    lrs!edicao = Now
    
    lrs.Update
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
    
    result = MsgBox("CADASTRO ATUALIZADO COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA LIMPAR OS DADOS?", vbYesNo + vbInformation)
    If result = vbYes Then
        CommandButton102_Click
        UserForm1.MultiPage1.Value = 0
    End If
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub CommandButton3_Click()
'OK BAIXA
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR REALIZAR UMA PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    If OptionButton4 = True Then
        MsgBox "NÃO É POSSÍVEL BAIXAR UM ITEM JÁ BAIXADO !" & vbCrLf & vbCrLf & "FAVOR ALTERAR O TIPO DE PESQUISA", vbCritical
        GoTo FIM
    ElseIf OptionButton5 = True Then
        MsgBox "NÃO É POSSÍVEL BAIXAR UM ITEM EXCLUIDO !" & vbCrLf & vbCrLf & "FAVOR ALTERAR O TIPO DE PESQUISA", vbCritical
        GoTo FIM
    End If
    
    ID = ListBox1.List(ListBox1.ListIndex, 0)
    If ID = 0 Or Not IsNumeric(ID) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA DELETAR", vbCritical
        GoTo FIM
    End If
    
    QTD = Application.InputBox("INFORME A QUANTIDADE QUE DESEJA DAR BAIXA" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA")
    If QTD > 0 Then
        Set ws = Sheets("Base de dados")
        Application.ScreenUpdating = False
        If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
        On Error Resume Next
    
        ws.ShowAllData
        Set rngAF = ws.Range("A1:A" & ws.Cells(1, 1).End(xlDown).Row)
        ws.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
    
        lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row

        If ws.Cells(lin_inicio, 5).Value - CInt(QTD) > 0 Then
            lsConectar
            Set lrs = New ADODB.Recordset
            lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
            lrs!qt_pecas = ws.Cells(lin_inicio, 8).Value - QTD
            lrs.Update
            Set lrs = Nothing
            
            Set lrs = New ADODB.Recordset
            sql = " INSERT INTO Baixa "
            sql = sql & " (item, descricao, qt_pecas, area_estoque, comentario, entrada, edicao, saida) "
            sql = sql & " VALUES "
            sql = sql & " ('" & ws.Cells(lin_inicio, 2) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 3) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 4) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 5) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 6) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 7) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 8) & "', "
            sql = sql & " '" & Now & "') "
            
            lrs.Open sql, gConexao
            lrs.Close
            Set lrs = Nothing
            lsDesconectar
            
        ElseIf ws.Cells(lin_inicio, 5).Value - CInt(QTD) = 0 Then
            lsConectar
            Set lrs = New ADODB.Recordset
            lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
            lrs.Delete
            lrs.Update
            Set lrs = Nothing
            
            Set lrs = New ADODB.Recordset
            sql = " INSERT INTO Baixa "
            sql = sql & " (item, descricao, qt_pecas, area_estoque, comentario, entrada, edicao, saida) "
            sql = sql & " VALUES "
            sql = sql & " ('" & ws.Cells(lin_inicio, 2) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 3) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 4) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 5) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 6) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 7) & "', "
            sql = sql & " '" & ws.Cells(lin_inicio, 8) & "', "
            sql = sql & " '" & Now & "') "
            
            lrs.Open sql, gConexao
            lrs.Close
            Set lrs = Nothing
            lsDesconectar
        Else
            MsgBox "VALOR INFORMADO MAIOR QUE ESTOQUE" & vbCrLf & vbCrLf & "VERIFIQUE A QUANTIDADE INFORMADA"
            GoTo FIM
        End If
        
        ws.ShowAllData
        result = MsgBox("BAIXA REALIZADA COM SUCESSO!", vbOKOnly + vbInformation)
        OptionButton3_Click
    End If
FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextBox3.MaxLength = 8
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        If TextBox3.SelStart = 2 Then
            TextBox3.SelText = "-"
        End If
        If TextBox3.SelStart = 5 Then
            TextBox3.SelText = "-"
        End If
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox102_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If TextBox102 <> "" And TextBox101 = "" Then
        Sheets("Descrição").Cells(1, 4) = TextBox102.Text
        TextBox103 = UCase(Worksheets("Descrição").Cells(1, 5).Value)
    End If
End Sub

Private Sub TextBox102_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox103_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox104_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox105_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextBox105.MaxLength = 8
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        If TextBox105.SelStart = 2 Then
            TextBox105.SelText = "-"
        End If
        If TextBox105.SelStart = 5 Then
            TextBox105.SelText = "-"
        End If
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox106_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
