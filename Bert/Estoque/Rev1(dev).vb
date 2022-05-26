'Microsoft ActiveX Data Objects 6.1 Library
'Option Explicit
Dim gConexao As New ADODB.Connection
Dim lrs As New ADODB.Recordset
Dim strConexao, sql As String
Dim ws, wsB, wsC As Worksheet
Dim wb As Workbook

Private Sub lsConectar()
    Set gConexao = New ADODB.Connection
    
    'strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\Grupos\CZ1 - Transferencia Informacoes\10. Métodos e Processos\BD_Exp_BSA\Database_EXP_BSA.accdb;Persist Security Info=False"
    strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\willian.cardoso\Desktop\TESTE.accdb;Persist Security Info=False"
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
    Sheets("PRJ").Visible = xlSheetVeryHidden
    Sheets("Componentes").Visible = xlSheetVeryHidden
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
        Sheets("PRJ").Visible = xlSheetVisible
        Sheets("Componentes").Visible = xlSheetVisible
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

Private Sub CommandButton1_Click()
'OK PESQUISA'
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    Set ws = Sheets("Base de dados")
    
    If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    ws.ShowAllData
    ws.Rows("2:" & ws.Cells(1, 1).End(xlDown).Row).ClearContents
    
    lsConectar
    Set lrs = New ADODB.Recordset

'''''''SELEÇÃO DO BANCO DE DADOS FILTRADO'''''''
'    If TextBox3 <> "" Then lrs.Open " SELECT * FROM BD_dados WHERE pedido = '" & TextBox3 & "'", gConexao, adOpenKeyset, adLockPessimistic

    lrs.Open " SELECT * FROM BD_dados ", gConexao, adOpenKeyset, adLockPessimistic
    ws.Cells(2, 1).CopyFromRecordset lrs
    ws.Columns("A:A").NumberFormat = "0"
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
    
    Set rngAF = ws.Range("A1:A" & ws.Cells(1, 1).End(xlDown).Row)
    
    If TextBox1 <> "" Then ws.Range("B:B").AutoFilter Field:=2, Criteria1:="=*" & TextBox1.Text & "*"
    If TextBox2 <> "" Then ws.Range("C:C").AutoFilter Field:=3, Criteria1:="=*" & TextBox2.Text & "*"
    If TextBox3 <> "" Then ws.Range("D:D").AutoFilter Field:=4, Criteria1:=TextBox3.Text
    If TextBox4 <> "" Then ws.Range("F:F").AutoFilter Field:=6, Criteria1:=TextBox4.Text
    If TextBox5 <> "" Then ws.Range("G:G").AutoFilter Field:=7, Criteria1:="=*" & TextBox5.Text & "*"
    If TextBox6 <> "" Then ws.Range("E:E").AutoFilter Field:=5, Criteria1:=TextBox6.Text
    If TextBox7 <> "" Then ws.Range("I:I").AutoFilter Field:=9, Criteria1:=TextBox7.Text
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = ws.Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If ws.Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        ListBox1.Clear
        ws.ShowAllData
        GoTo FIM
    End If
    
    Dim arrayItems2()
    With Planilha5
        ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, ws.Range("A:A")), 1 To 11)
        Me.ListBox1.ColumnCount = 11
        Me.ListBox1.ColumnWidths = "40;120;300;70;70;70;200;70;70;70;200"
        i = 0
        For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
            Me.ListBox1.AddItem
            For coluna = 1 To 11
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
            .FitToPagesTall = False
            .FitToPagesWide = 1
            .CenterHorizontally = True
            .PaperSize = xlPaperA4
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
End Sub

Private Sub CommandButton5_Click()
'APAGAR PESQUISA
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Set ws = Sheets("Base de dados")
    Set wsB = Sheets("Baixados")

    ID = ListBox1.List(ListBox1.ListIndex, 0)
    If ID = 0 Or Not IsNumeric(ID) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA DELETAR", vbCritical
        GoTo FIM
    End If
    
    result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR O ID " & ID & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
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
    CheckBox1 = False
    CheckBox2 = False
    ListBox1.Clear
End Sub

Private Sub CommandButton4_Click()
'EDITAR PESQUISA

    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    On Error Resume Next
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
        TextBox102.SetFocus
    End If
    
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub TextBox102_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'CADASTRO - IMPORTAÇÃO DADOS ATRAVES O.P.

    If TextBox101 = "" And TextBox102 <> "" Then
        Set wsC = Sheets("Componentes")
        Application.ScreenUpdating = False
        
        On Error Resume Next
        wsC.ListObjects(1).ShowAutoFilter = True
        wsC.ListObjects(1).AutoFilter.ShowAllData
        
        wsC.ListObjects(1).Range.AutoFilter Field:=5, Criteria1:=TextBox102
        
        lin_inicio = wsC.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        If wsC.Cells(lin_inicio, 1).Value = 0 Then
            wsC.ListObjects(1).AutoFilter.ShowAllData
            TextBox103 = ""
            TextBox104 = ""
            TextBox105 = ""
            TextBox106 = ""
            TextBox107 = ""
            GoTo FIM
        End If
        
        TextBox103 = wsC.Cells(lin_inicio, 2)
        TextBox104 = wsC.Cells(lin_inicio, 3)
        TextBox105 = wsC.Cells(lin_inicio, 4)
        TextBox106 = wsC.Cells(lin_inicio, 6)
        TextBox107 = wsC.Cells(lin_inicio, 7)
        
        wsC.ListObjects(1).AutoFilter.ShowAllData
        
FIM:
        Application.ScreenUpdating = True
    ElseIf TextBox101 = "" And TextBox102 = "" Then
        TextBox103 = ""
        TextBox104 = ""
        TextBox105 = ""
        TextBox106 = ""
        TextBox107 = ""
    End If
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
'OK BAIXA
    
    If OptionButton3.Value = True And TextBox201 = 0 And TextBox203 <> "" And TextBox204 <> "" Then
    ElseIf TextBox201 <> "" And TextBox202 <> "" And (OptionButton1.Value = True Or OptionButton2.Value = True) And TextBox204 <> "" Then
    ElseIf OptionButton3.Value = True And TextBox201 <> 0 Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "PEDIDO INFORMADO NÃO É KANBAN", vbCritical
        GoTo FIM
    Else
        MsgBox "FAVOR PREENCHER AS INFORMAÇÕES PARA CONTINUAR!", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Set ws = Sheets("Base de dados")
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

                Set lrs = New ADODB.Recordset
                lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
                lrs.Delete
                lrs.Update
                Set lrs = Nothing
                
                Set lrs = New ADODB.Recordset
                sql = " INSERT INTO Baixa "
                sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada, edicao, saida, cod_coleta) "
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
                sql = sql & " '" & Now & "', "
                sql = sql & " '" & TextBox204 & "') "
                
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
        TextBox204 = ""
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
            Set lrs = New ADODB.Recordset
            lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
            lrs.Delete
            lrs.Update
            Set lrs = Nothing
            
            Set lrs = New ADODB.Recordset
            sql = " INSERT INTO Baixa "
            sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada, edicao, saida, cod_coleta) "
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
            sql = sql & " '" & Now & "', "
            sql = sql & " '" & TextBox204 & "') "
            
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
                TextBox204 = ""
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
                Set lrs = New ADODB.Recordset
                lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
                lrs!qt_pecas = ws.Cells(lin_inicio, 8).Value - QTD
                lrs.Update
                Set lrs = Nothing
                
                Set lrs = New ADODB.Recordset
                sql = " INSERT INTO Baixa "
                sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada, edicao, saida, cod_coleta) "
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
                sql = sql & " '" & Now & "', "
                sql = sql & " '" & TextBox204 & "') "
                
                lrs.Open sql, gConexao
                lrs.Close
                Set lrs = Nothing
                lsDesconectar
                
            ElseIf ws.Cells(lin_inicio, 8).Value - QTD = 0 Then
                
                lsConectar
                Set lrs = New ADODB.Recordset
                lrs.Open "SELECT * FROM BD_dados WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
                lrs.Delete
                lrs.Update
                Set lrs = Nothing
                
                Set lrs = New ADODB.Recordset
                sql = " INSERT INTO Baixa "
                sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada, edicao, saida, cod_coleta) "
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
                sql = sql & " '" & Now & "', "
                sql = sql & " '" & TextBox204 & "') "
                
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
                TextBox204 = ""
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

Private Sub CommandButton202_Click()
'LIMPAR BAIXA

    TextBox201 = ""
    TextBox202 = ""
    TextBox203 = ""
    TextBox204 = ""
    OptionButton1.Value = False
    OptionButton2.Value = False
    OptionButton3.Value = False
    ListBox2.Clear
End Sub

Private Sub CommandButton203_Click()
'NOVO CÓD COLETA
    lsConectar
    Set lrs = New ADODB.Recordset
    lrs.Open " SELECT MAX(cod_coleta) FROM Baixa ", gConexao, adOpenKeyset, adLockPessimistic
    TextBox204 = CInt(lrs.GetString) + 1
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
End Sub

Private Sub CommandButton204_Click()
'ULTIMO CÓD COLETA
    lsConectar
    Set lrs = New ADODB.Recordset
    lrs.Open " SELECT MAX(cod_coleta) FROM Baixa ", gConexao, adOpenKeyset, adLockPessimistic
    TextBox204 = CInt(lrs.GetString)
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
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
    TextBox303 = ListBox3.List(ListBox3.ListIndex, 0)
    TextBox304 = ListBox3.List(ListBox3.ListIndex, 4)
End Sub

Private Sub CommandButton401_Click()
'OK RELATÓRIO BAIXA
Application.ScreenUpdating = False
    Set wb = ThisWorkbook
    Set wsB = Sheets("Baixados")
    
    If Not wsB.AutoFilterMode Then wsB.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    wsB.ShowAllData
    wsB.Rows("2:" & wsB.Cells(1, 1).End(xlDown).Row).ClearContents
    
    lsConectar
    Set lrs = New ADODB.Recordset
    lrs.Open " SELECT * FROM Baixa ", gConexao, adOpenKeyset, adLockPessimistic
    wsB.Cells(2, 1).CopyFromRecordset lrs
    wsB.Columns("A:A").NumberFormat = "0"
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
    
    Set rngAF = wsB.Range("A1:A" & wsB.Cells(1, 1).End(xlDown).Row)
    
    If TextBox401 <> "" Then wsB.Range("O:O").AutoFilter Field:=15, Criteria1:=TextBox401
    If TextBox402 <> "" Then wsB.Range("D:D").AutoFilter Field:=4, Criteria1:=TextBox402
    If TextBox403 <> "" Then wsB.Range("G:G").AutoFilter Field:=7, Criteria1:="=*" & TextBox403 & "*"
    
    lin_inicio = wsB.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = wsB.Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If wsB.Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        ListBox4.Clear
        wsB.ShowAllData
        GoTo FIM
    End If
    
    Dim arrayItems2()
    With Planilha3
        ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, wsB.Range("A:A")), 1 To 15)
        Me.ListBox4.ColumnCount = 15
        Me.ListBox4.ColumnWidths = "40;120;300;70;70;70;200;70;0;0;0;70;0;70;"
        i = 0
        For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
            Me.ListBox4.AddItem
            For coluna = 1 To 15
                arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
            Next coluna
            i = i + 1
        Next rngcell
        Me.ListBox4.List = arrayItems2()
    End With
    
    If CheckBox3 = True Or CheckBox4 = True Then
        Set rngAJ = wsB.Range("A1:O" & lin_fim).SpecialCells(xlCellTypeVisible)
        rngAJ.Copy
        Workbooks.Add
        Range("A1").PasteSpecial Paste:=xlPasteValues
        Columns("A:A,I:M").Delete
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:I" & Cells(1, 1).End(xlDown).Row), , xlYes).Name = "Tabela1"
        Columns("A:I").EntireColumn.AutoFit
        
        
        
        
        'Cells(Cells(1, 1).End(xlDown).Row + 3, 1) = "Declaro para os devidos fins que o material recebido foi conferido e encontra-se em perfeitas condições"
        
        
        
        
        
        
        If CheckBox4 = True Then
            With Worksheets(1).PageSetup
            .Zoom = False
            .BlackAndWhite = True
            .FitToPagesTall = False
            .FitToPagesWide = 1
            .CenterHorizontally = True
            .PaperSize = xlPaperA4
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
        If CheckBox3 = False Then ActiveWorkbook.Close SaveChanges:=False
    End If
    If CheckBox3 = True Then
        result = MsgBox("DADOS EXPORTADOS COM SUCESSO !" & vbCrLf & "DESEJA FECHAR O FORMULÁRIO ?" & vbCrLf & vbCrLf & "É NECESSÁRIO FECHAR PARA EDITAR OS DADOS", vbYesNo + vbInformation)
        If result = vbYes Then
            Unload Me
        Else
            ActiveWindow.WindowState = xlMinimized
            wb.Activate
            CheckBox3 = False
            CheckBox4 = False
        End If
    End If
FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton403_Click()
'EDITAR RELATÓRIO BAIXA

    If ListBox4.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Set wsB = Sheets("Baixados")
    
    ID = ListBox4.List(ListBox4.ListIndex, 0)
    If ID = 0 Or Not IsNumeric(ID) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA EDITAR", vbCritical
        GoTo FIM
    End If
    
    If Not wsB.AutoFilterMode Then wsB.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    wsB.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
    
    lin_inicio = wsB.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    UserForm2.TextBox1 = wsB.Cells(lin_inicio, 15)
    UserForm2.TextBox2 = wsB.Cells(lin_inicio, 1)
    UserForm2.TextBox3 = wsB.Cells(lin_inicio, 2)
    UserForm2.TextBox4 = wsB.Cells(lin_inicio, 3)
    UserForm2.TextBox5 = wsB.Cells(lin_inicio, 4)
    UserForm2.TextBox6 = wsB.Cells(lin_inicio, 7)
    UserForm2.TextBox7 = wsB.Cells(lin_inicio, 8)
    
    wsB.ShowAllData
    ListBox4.Clear
    UserForm2.Show
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub CommandButton404_Click()
'APAGAR RELATÓRIO BAIXA
    If ListBox4.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Set ws = Sheets("Base de dados")
    Set wsB = Sheets("Baixados")
    
    ID = ListBox4.List(ListBox4.ListIndex, 0)
    If ID = 0 Or Not IsNumeric(ID) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA EDITAR", vbCritical
        GoTo FIM
    End If
        
    result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR O ID " & ID & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
    If result = vbYes Then
    
        Application.ScreenUpdating = False
        If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
        On Error Resume Next
        wsB.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
        
        lin_inicio = wsB.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        lsConectar
        Set lrs = New ADODB.Recordset
        lrs.Open "SELECT * FROM Baixa WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
        lrs.Delete
        lrs.Update
        Set lrs = Nothing
        
        Set lrs = New ADODB.Recordset
        sql = " INSERT INTO BD_dados "
        sql = sql & " (item, descricao, pedido, ordem_prod, programa, cliente, qt_pecas, area_estoque, posicao, comentario, entrada, edicao) "
        sql = sql & " VALUES "
        sql = sql & " ('" & wsB.Cells(lin_inicio, 2) & "', "
        sql = sql & " '" & wsB.Cells(lin_inicio, 3) & "', "
        sql = sql & " '" & wsB.Cells(lin_inicio, 4) & "', "
        sql = sql & " '" & wsB.Cells(lin_inicio, 5) & "', "
        sql = sql & " '" & wsB.Cells(lin_inicio, 6) & "', "
        sql = sql & " '" & wsB.Cells(lin_inicio, 7) & "', "
        sql = sql & " '" & wsB.Cells(lin_inicio, 8) & "', "
        sql = sql & " '" & wsB.Cells(lin_inicio, 9) & "', "
        sql = sql & " '" & wsB.Cells(lin_inicio, 10) & "', "
        sql = sql & " '" & wsB.Cells(lin_inicio, 11) & "', "
        sql = sql & " '" & wsB.Cells(lin_inicio, 12) & "', "
        sql = sql & " '" & wsB.Cells(lin_inicio, 13) & "') "
        
        lrs.Open sql, gConexao
        lrs.Close
        Set lrs = Nothing
        lsDesconectar
        
        MsgBox "CADASTRO EXCLUIDO COM SUCESSO!", vbInformation
        ListBox4.Clear
    End If
FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton402_Click()
'LIMPAR RELATÓRIO BAIXA
    TextBox401 = ""
    TextBox402 = ""
    TextBox403 = ""
    CheckBox3 = False
    CheckBox4 = False
    ListBox4.Clear
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

Private Sub TextBox4021_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextBox4021.MaxLength = 8
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        If TextBox4021.SelStart = 2 Then
            TextBox4021.SelText = "/"
        End If
        If TextBox4021.SelStart = 5 Then
            TextBox4021.SelText = "/"
        End If
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox4022_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextBox4022.MaxLength = 8
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        If TextBox4022.SelStart = 2 Then
            TextBox4022.SelText = "/"
        End If
        If TextBox4022.SelStart = 5 Then
            TextBox4022.SelText = "/"
        End If
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Option Explicit
Dim gConexao As New ADODB.Connection
Dim lrs As New ADODB.Recordset
Dim strConexao, sql As String
Dim ws, wsB, wsC As Worksheet
Dim wb As Workbook

Private Sub lsConectar()
    Set gConexao = New ADODB.Connection
    
    'strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\Grupos\CZ1 - Transferencia Informacoes\10. Métodos e Processos\BD_Exp_BSA\Database_EXP_BSA.accdb;Persist Security Info=False"
    strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\willian.cardoso\Desktop\TESTE.accdb;Persist Security Info=False"
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
Private Sub CommandButton102_Click()
'CANCELAR
    Unload Me
End Sub

Private Sub CommandButton104_Click()
'ATUALIZAR
    
    If UserForm2.TextBox1 = "" Then
        MsgBox "FAVOR PREENCHER O CÓD DE COLETA!", vbCritical
        GoTo FIM
    End If
    
    lsConectar
    Set lrs = New ADODB.Recordset
    lrs.Open "SELECT * FROM Baixa WHERE ID = " & CInt(UserForm2.TextBox2), gConexao, adOpenKeyset, adLockPessimistic
    
    lrs!cod_coleta = UserForm2.TextBox1
    
    lrs.Update
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
    
    MsgBox "CADASTRO ATUALIZADO COM SUCESSO!", vbInformation
    Unload Me
FIM:
End Sub
