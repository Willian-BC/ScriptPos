'Option Explicit
Dim gConexao As New ADODB.Connection
Dim lrs As New ADODB.Recordset
Dim strConexao, sql As String
Dim ws, wsB, wsC As Worksheet
Dim wb As Workbook

Private Sub lsConectar()
    Set gConexao = New ADODB.Connection
    
'    strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\Grupos\CZ1 - Transferencia Informacoes\10. Métodos e Processos\BD_Exp_BSA\Database_EXP_BSA.accdb;Persist Security Info=False"
    strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\Grupos\CZ1 - Transferencia Informacoes\10. Métodos e Processos\Database_EXP_BSA.accdb;Persist Security Info=False"
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

Public Sub UserForm_Initialize()
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    Sheets("Baixados").Visible = xlSheetVeryHidden
    Sheets("Componentes").Visible = xlSheetVeryHidden
    Sheets("Componentes (2)").Visible = xlSheetVeryHidden
    If (Now - Sheets("Planilha1").Cells(3, 3)) < TimeValue("00:00:15") Then
        Sheets("Planilha1").Range("G2") = TimeValue("00:00:15") - (Now - Sheets("Planilha1").Cells(3, 3))
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
    If MyValue = "1010" Then
        Sheets("Componentes").Visible = xlSheetVisible
    ElseIf MyValue = "963" Then
        Sheets("Base de dados").Visible = xlSheetVisible
        Sheets("Baixados").Visible = xlSheetVisible
        Sheets("Componentes").Visible = xlSheetVisible
        Sheets("Componentes (2)").Visible = xlSheetVisible
    Else
        MsgBox ("Senha Incorreta")
    End If
End Sub

Private Sub CommandButton200_Click()
    ThisWorkbook.Save
End Sub

Private Sub CommandButton300_Click()
    On Error Resume Next
    Application.ScreenUpdating = False
    If Not Sheets("Base de dados").AutoFilterMode Then Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
    Sheets("Base de dados").ShowAllData
    Worksheets("Componentes").ListObjects(1).ShowAutoFilter = True
    Worksheets("Componentes").ListObjects(1).AutoFilter.ShowAllData
    Application.ScreenUpdating = True
End Sub

Private Sub OptionButton2_Click()
    ListBox1.Clear
End Sub
Private Sub OptionButton3_Click()
    ListBox1.Clear
End Sub

Public Sub CommandButton1_Click()
'OK PESQUISA'
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlManual
    
    Set wb = ActiveWorkbook
    If OptionButton2 = True Then
        Set ws = Sheets("Base de dados")
        a = 12
        b = "L:L"
        If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
        On Error Resume Next
        ws.ShowAllData
        ws.Rows("2:" & ws.Cells(1, 1).End(xlDown).Row).ClearContents
        
        lsConectar
        Set lrs = New ADODB.Recordset
'        lrs.Open " SELECT * FROM BD_dados ", gConexao, adOpenKeyset, adLockPessimistic
        If TextBox3 <> "" Then
            lrs.Open " SELECT * FROM BD_dados WHERE pedido = '" & TextBox3 & "'", gConexao, adOpenKeyset, adLockPessimistic
        ElseIf TextBox6 <> "" Then
            lrs.Open " SELECT * FROM BD_dados WHERE ordem_prod = '" & TextBox6 & "'", gConexao, adOpenKeyset, adLockPessimistic
        Else
            lrs.Open " SELECT * FROM BD_dados ", gConexao, adOpenKeyset, adLockPessimistic
        End If
        
        ws.Cells(2, 1).CopyFromRecordset lrs
        ws.Columns("A:A").NumberFormat = "0"
        lrs.Close
        Set lrs = Nothing
        lsDesconectar
    ElseIf OptionButton3 = True Then
        Set ws = Sheets("Baixados")
        a = 14
        b = "N:N"
        If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
        On Error Resume Next
        ws.ShowAllData
        ws.Rows("2:" & ws.Cells(1, 1).End(xlDown).Row).ClearContents
        
        lsConectar
        Set lrs = New ADODB.Recordset
'        lrs.Open " SELECT * FROM Baixa ", gConexao, adOpenKeyset, adLockPessimistic
        If TextBox3 <> "" Then
            lrs.Open " SELECT * FROM Baixa WHERE pedido = '" & TextBox3 & "'", gConexao, adOpenKeyset, adLockPessimistic
        ElseIf TextBox6 <> "" Then
            lrs.Open " SELECT * FROM Baixa WHERE ordem_prod = '" & TextBox6 & "'", gConexao, adOpenKeyset, adLockPessimistic
        Else
            lrs.Open " SELECT * FROM Baixa ", gConexao, adOpenKeyset, adLockPessimistic
        End If
        
        ws.Cells(2, 1).CopyFromRecordset lrs
        ws.Columns("A:A").NumberFormat = "0"
        lrs.Close
        Set lrs = Nothing
        lsDesconectar
    End If
    
    Set rngAF = ws.Range("A1:A" & ws.Cells(1, 1).End(xlDown).Row)
    
    If TextBox1 <> "" Then ws.Range("B:B").AutoFilter Field:=2, Criteria1:="=*" & TextBox1.Text & "*"
    If TextBox2 <> "" Then ws.Range("C:C").AutoFilter Field:=3, Criteria1:="=*" & TextBox2.Text & "*"
    If TextBox3 <> "" Then ws.Range("D:D").AutoFilter Field:=4, Criteria1:=TextBox3.Text
    If TextBox4 <> "" Then ws.Range("J:J").AutoFilter Field:=10, Criteria1:=TextBox4.Text
    If TextBox5 <> "" Then ws.Range("G:G").AutoFilter Field:=7, Criteria1:="=*" & TextBox5.Text & "*"
    If TextBox6 <> "" Then ws.Range("E:E").AutoFilter Field:=5, Criteria1:=TextBox6.Text
    If TextBox7 <> "" Then ws.Range("I:I").AutoFilter Field:=9, Criteria1:=TextBox7.Text
    If TextBox8 <> "" Then ws.Range(b).AutoFilter Field:=a, Criteria1:="=*" & Format(CDate(TextBox8), "dd/mm/yyyy") & "*"
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = ws.Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If ws.Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        ListBox1.Clear
        ws.ShowAllData
        GoTo FIM
    End If
    
    Dim arrayItems2()
    With ws
        ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, ws.Range("A:A")), 1 To a)
        Me.ListBox1.ColumnCount = a
        Me.ListBox1.ColumnWidths = "40;120;300;70;70;70;200;70;70;70;200"
        i = 0
        For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
            Me.ListBox1.AddItem
            For coluna = 1 To a
                arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
            Next coluna
            i = i + 1
        Next rngcell
        Me.ListBox1.List = arrayItems2()
    End With
    
    If CheckBox1 = True Or CheckBox2 = True Then
        Set rngAJ = ws.Range("A1:K" & lin_fim).SpecialCells(xlCellTypeVisible)
        rngAJ.Copy
        Workbooks.Add
        Range("A1").PasteSpecial Paste:=xlPasteValues
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:K" & Cells(1, 1).End(xlDown).Row), , xlYes).Name = "Tabela1"
        Columns("A:K").EntireColumn.AutoFit
        If CheckBox2 = True Then
            With Worksheets(1).PageSetup
            .Zoom = False
            .BlackAndWhite = True
            .FitToPagesTall = 1
            .FitToPagesWide = 1
            .CenterHorizontally = True
            .Orientation = xlLandscape
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            End With
            ActiveSheet.PrintOut
            MsgBox "IMPRESSÃO REALIZADA COM SUCESSO!", vbInformation
        End If
        If CheckBox1 = False Then ActiveWorkbook.Close SaveChanges:=False
    End If
    If CheckBox1 = True Then
        result = MsgBox("DADOS EXPORTADOS COM SUCESSO !" & vbCrLf & "DESEJA FECHAR O FORMULÁRIO ?" & vbCrLf & vbCrLf & "É NECESSÁRIO FECHAR PARA EDITAR OS DADOS", vbYesNo + vbInformation)
        If result = vbYes Then
            Unload Me
        Else
            ActiveWindow.WindowState = xlMinimized
            wb.Activate
            CheckBox1 = False
            CheckBox2 = False
        End If
    End If
FIM:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlAutomatic
End Sub

Private Sub CommandButton5_Click()
'APAGAR PESQUISA
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    If OptionButton3 = True Then
        MsgBox "NÃO É POSSÍVEL DELETAR UM ITEM BAIXADO !" & vbCrLf & vbCrLf & "FAVOR ALTERAR O TIPO DE PESQUISA", vbCritical
        GoTo FIM
    End If
    
    Set ws = Sheets("Base de dados")
    
    If ListBox1.ListIndex = -1 Or ListBox1.ListIndex = 0 Or Not IsNumeric(ListBox1.ListIndex) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA EDITAR", vbCritical
        GoTo FIM
    End If
    ID = ListBox1.List(ListBox1.ListIndex, 0)
    
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
FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton2_Click()
'LIMPAR PESQUISA
    TextBox1 = ""
    TextBox2 = ""
    TextBox3 = ""
    TextBox4 = ""
    TextBox5 = ""
    TextBox6 = ""
    TextBox7 = ""
    TextBox8 = ""
    CheckBox1 = False
    CheckBox2 = False
    OptionButton2 = True
    OptionButton3 = False
    ListBox1.Clear
End Sub

Private Sub CommandButton4_Click()
'EDITAR PESQUISA
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    If OptionButton3 = True Then
        MsgBox "NÃO É POSSÍVEL EDITAR UM ITEM BAIXADO !" & vbCrLf & vbCrLf & "FAVOR ALTERAR O TIPO DE PESQUISA", vbCritical
        GoTo FIM
    End If
    
    Set ws = Sheets("Base de dados")
    
    If ListBox1.ListIndex = -1 Or ListBox1.ListIndex = 0 Or Not IsNumeric(ListBox1.ListIndex) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA EDITAR", vbCritical
        GoTo FIM
    End If
    ID = ListBox1.List(ListBox1.ListIndex, 0)
    
    Application.ScreenUpdating = False
    
    If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    ws.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    TextBox101 = ws.Cells(lin_inicio, 1)
    TextBox102 = ws.Cells(lin_inicio, 5)
    TextBox103 = ws.Cells(lin_inicio, 2)
    TextBox104 = ws.Cells(lin_inicio, 3)
    TextBox105 = ws.Cells(lin_inicio, 4)
    TextBox106 = ws.Cells(lin_inicio, 6)
    TextBox107 = ws.Cells(lin_inicio, 7)
    TextBox108 = ws.Cells(lin_inicio, 8)
    TextBox109 = ws.Cells(lin_inicio, 9)
    TextBox110 = ws.Cells(lin_inicio, 10)
    TextBox111 = ws.Cells(lin_inicio, 11)
    
    ws.ShowAllData
    ListBox1.Clear
    UserForm1.MultiPage1.Value = 1
    Application.ScreenUpdating = True
FIM:
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
    
    lrs!Item = TextBox103
    lrs!descricao = TextBox104
    lrs!pedido = TextBox105
    lrs!ordem_prod = TextBox102
    lrs!programa = TextBox106
    lrs!cliente = TextBox107
    lrs!qt_pecas = TextBox108
    lrs!area_estoque = TextBox109
    lrs!posicao = TextBox110
    lrs!comentario = TextBox111
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

Private Sub CommandButton101_Click()
'OK CADASTRO'
    If TextBox101 <> "" Then
        MsgBox "CADASTRO JÁ EXISTE!" & vbCrLf & vbCrLf & "CLIQUE EM ATUALIZAR REGISTRO OU LIMPAR", vbCritical
        GoTo FIM
    ElseIf TextBox102 = "" Or TextBox103 = "" Or TextBox105 = "" Or TextBox108 = "" Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS PARA CADASTRO!", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    
    lsConectar
    Set lrs = New ADODB.Recordset
    
    sql = " INSERT INTO BD_dados "
    sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada) "
    sql = sql & " VALUES "
    sql = sql & " ('" & TextBox103 & "', "
    sql = sql & " '" & TextBox104 & "', "
    sql = sql & " '" & TextBox105 & "', "
    sql = sql & " '" & TextBox102 & "', "
    sql = sql & " '" & TextBox106 & "', "
    sql = sql & " '" & TextBox107 & "', "
    sql = sql & " '" & TextBox108 & "', "
    sql = sql & " '" & TextBox109 & "', "
    sql = sql & " '" & TextBox110 & "', "
    sql = sql & " '" & TextBox111 & "', "
    sql = sql & " '" & Now & "') "
    
    lrs.Open sql, gConexao
    Set lrs = Nothing
    lsDesconectar
    
    result = MsgBox("CADASTRO REALIZADO COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA LIMPAR OS DADOS?", vbYesNo + vbInformation)
    If result = vbYes Then
        CommandButton102_Click
        TextBox102.SetFocus
    End If
    
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub TextBox102_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'CADASTRO - IMPORTAÇÃO DADOS ATRAVES O.P.

    If TextBox101 = "" And TextBox102 <> "" Then
        Set wsC = Sheets("Componentes (2)")
        Application.ScreenUpdating = False
        
        On Error Resume Next
        wsC.ListObjects(1).ShowAutoFilter = True
        wsC.ListObjects(1).AutoFilter.ShowAllData
        
        wsC.ListObjects(1).Range.AutoFilter Field:=5, Criteria1:=TextBox102
        
        lin_inicio = wsC.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        If wsC.Cells(lin_inicio, 1).Value = 0 Then
            wsC.ListObjects(1).AutoFilter.ShowAllData
            a = TextBox102
            CommandButton102_Click
            TextBox102 = a
            GoTo FIM
        End If
        
        TextBox103 = wsC.Cells(lin_inicio, 2)
        TextBox104 = wsC.Cells(lin_inicio, 3)
        TextBox105 = wsC.Cells(lin_inicio, 4)
        TextBox106 = wsC.Cells(lin_inicio, 6)
        TextBox107 = wsC.Cells(lin_inicio, 7)
        
        wsC.ListObjects(1).AutoFilter.ShowAllData
        
        If Right(Left(TextBox103, Len(TextBox103) - 2), 1) <> 0 Then
            TextBox110 = "SAX11"
        Else
            TextBox110 = ""
        End If
        
    ElseIf TextBox101 = "" And TextBox102 = "" Then
        a = TextBox102
        CommandButton102_Click
        TextBox102 = a
    End If
FIM:
        Application.ScreenUpdating = True
End Sub

Private Sub CommandButton102_Click()
'LIMPAR CADASTRO
    TextBox101 = ""
    TextBox102 = ""
    TextBox103 = ""
    TextBox104 = ""
    TextBox105 = ""
    TextBox106 = ""
    TextBox107 = ""
    TextBox108 = ""
    TextBox109 = ""
    TextBox110 = ""
    TextBox111 = ""
End Sub

Private Sub CommandButton201_Click()
'PESQUISAR BAIXA
    If TextBox201 = "" Then
        MsgBox "FAVOR INFORMAR UM PEDIDO PARA CONTINUAR", vbCritical
        GoTo FIM
    ElseIf TextBox201 = 0 And (OptionButton1 = False Or TextBox202 = "") Then
        MsgBox "FAVOR PREENCHER O CÓDIGO ITEM E MARCAR OPÇÃO KANBAN PARA CONTINUAR!", vbCritical
        GoTo FIM
    ElseIf TextBox201 <> 0 And OptionButton1 = True Then
        MsgBox "PEDIDO NÃO É KANBAN!", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Set ws = Sheets("Base de dados")
    ListBox1.Clear
    
    If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    ws.ShowAllData
    ws.Rows("2:" & ws.Cells(1, 1).End(xlDown).Row).ClearContents
    
    lsConectar
    Set lrs = New ADODB.Recordset
'    lrs.Open " SELECT * FROM BD_dados ", gConexao, adOpenKeyset, adLockPessimistic
    lrs.Open " SELECT * FROM BD_dados WHERE pedido = '" & TextBox201 & "'", gConexao, adOpenKeyset, adLockPessimistic
    ws.Cells(2, 1).CopyFromRecordset lrs
    ws.Columns("A:A").NumberFormat = "0"
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
    
    Set rngAF = ws.Range("A1:A" & ws.Cells(1, 1).End(xlDown).Row)
    
    If TextBox202 <> "" Then ws.Range("B:B").AutoFilter Field:=2, Criteria1:=TextBox202
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = ws.Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If ws.Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALOR INFORMADO NÃO EXISTE !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        ws.ShowAllData
        GoTo FIM
    End If
    
'    With Me.ListBox2
'        .List = ws.Range("A2:K" & lin_fim).SpecialCells(xlCellTypeVisible).Value
'        .ColumnHeads = True
'        .ColumnCount = 11
'        .ColumnWidths = "40;120;300;70;70;70;200;70;70;70;200"
'        .List(0, 1) = "Item"
'    End With
    
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
    TextBox203 = Format(WorksheetFunction.Subtotal(109, ws.Range("H:H")), "0")
FIM:
    Application.ScreenUpdating = True
End Sub
    
Private Sub CommandButton203_Click()
'BAIXAR BAIXA
    If ListBox2.ListCount = 0 Then
        MsgBox "FAVOR REALIZAR UMA PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Set ws = Sheets("Base de dados")
    Application.ScreenUpdating = False
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    If OptionButton1.Value = False Then
        ID = ListBox2.List(ListBox2.ListIndex, 0)
        If ID = 0 Or Not IsNumeric(ID) Then
            MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA EDITAR", vbCritical
            GoTo FIM
        End If
        ws.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
        lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        baixar (ID)
        ListBox2.RemoveItem (ListBox2.ListIndex)
        ws.Rows(lin_inicio).Delete
    ElseIf OptionButton1.Value = True Then
ERRO:
        QTD = Application.InputBox("INFORME A QUANTIDADE QUE DESEJA DAR BAIXAR")
        If QTD > 0 Then
            If TextBox203 - QTD >= 0 Then
                Do While QTD > 0
                    ID = ListBox2.List(1, 0)
                    
                    If ws.Cells(lin_inicio, 8).Value - QTD > 0 Then
                        lsConectar
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
'                        lrs.Close
                        Set lrs = Nothing
                        lsDesconectar
                        
                        Me.ListBox2.List(1, 7) = ws.Cells(lin_inicio, 8).Value - QTD
                        ws.Cells(lin_inicio, 8).Value = ws.Cells(lin_inicio, 8).Value - QTD
                        QTD = 0
                    Else
                        baixar (ID)
                        ListBox2.RemoveItem (1)
                        QTD = QTD - ws.Cells(lin_inicio, 8).Value
                        ws.Rows(lin_inicio).Delete
                        lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
                    End If
                Loop
            Else
                MsgBox "VALOR INFORMADO MAIOR QUE ESTOQUE" & vbCrLf & vbCrLf & "VERIFIQUE A QUANTIDADE INFORMADA"
                GoTo ERRO
            End If
        End If
    End If
    TextBox203 = Format(WorksheetFunction.Subtotal(109, ws.Range("H:H")), "0")
FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub baixar(ByVal ID As Integer)
    Set ws = Sheets("Base de dados")
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    lsConectar
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
'    lrs.Close
    Set lrs = Nothing
    lsDesconectar
End Sub

Private Sub CommandButton202_Click()
'LIMPAR BAIXA
    TextBox201 = ""
    TextBox202 = ""
    OptionButton1.Value = False
    ListBox2.Clear
End Sub

Private Sub CommandButton301_Click()
'OK APONTAMENTO - EXIBIR DADOS EM TELA

    If TextBox301 = "" Or TextBox302 = "" Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS!", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Set wsC = Sheets("Componentes")
    
    On Error Resume Next
    wsC.ListObjects(1).ShowAutoFilter = True
    wsC.ListObjects(1).AutoFilter.ShowAllData
    Set rngAF = wsC.Range("A1:A" & wsC.Cells(1, 1).End(xlDown).Row)
    
    wsC.ListObjects(1).Range.AutoFilter Field:=4, Criteria1:=TextBox301
    wsC.ListObjects(1).Range.AutoFilter Field:=3, Criteria1:="=*" & TextBox302 & "*"
    
    lin_inicio = wsC.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = wsC.Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If wsC.Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        wsC.ListObjects(1).AutoFilter.ShowAllData
        GoTo FIM
    End If
    
    Dim arrayItems2()
    With Planilha4
        ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, wsC.Range("A:A")), 1 To 10)
        Me.ListBox3.ColumnCount = 10
        Me.ListBox3.ColumnWidths = ";130;350;;;;200;;;"
        i = 0
        For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
            Me.ListBox3.AddItem
            For coluna = 1 To 10
                arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
            Next coluna
            i = i + 1
        Next rngcell
        Me.ListBox3.List = arrayItems2()
    End With
FIM:
End Sub

Private Sub CommandButton302_Click()
'LIMPAR APONTAMENTO

    TextBox301 = ""
    TextBox302 = ""
    TextBox303 = ""
    TextBox304 = ""
    ListBox3.Clear
End Sub

Private Sub TextBox303_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'APONTAMENTO - COLETAR O.P. ATRAVÉS DO INDICE

    If TextBox303 = "" Or ListBox3.ListCount = 0 Then
        TextBox304 = ""
    Else
        Application.ScreenUpdating = False
        Set wsC = Sheets("Componentes")
        
        wsC.ListObjects(1).Range.AutoFilter Field:=1, Criteria1:=TextBox303
        lin_inicio = wsC.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        If wsC.Cells(lin_inicio, 1).Value = 0 Then
            wsC.ListObjects(1).Range.AutoFilter Field:=1
            TextBox304 = ""
            GoTo FIM
        End If
        TextBox304 = wsC.Cells(lin_inicio, 5)
        TextBox304.SetFocus
    End If
FIM:
End Sub

Private Sub ListBox3_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ID = ListBox3.List(ListBox3.ListIndex, 0)
    If ID = 0 Or Not IsNumeric(ID) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO", vbCritical
        GoTo FIM
    End If
    TextBox304 = ListBox3.List(ListBox3.ListIndex, 4)
    TextBox304.SetFocus
FIM:
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox8_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextBox8.MaxLength = 10
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        If TextBox8.SelStart = 2 Then
            TextBox8.SelText = "/"
        End If
        If TextBox8.SelStart = 5 Then
            TextBox8.SelText = "/"
        End If
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox103_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox104_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox106_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox107_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox109_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox110_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox111_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox202_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox203_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox302_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
