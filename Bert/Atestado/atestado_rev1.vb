''''''''''''''''''''''''''''''''''''''''''EstaPastaDeTrabalho''''''''''''''''''''''''''''''''''''''''''
Private Sub Workbook_Open()
    Sheets("BD").Visible = xlSheetVeryHidden
    Sheets("Acesso").Visible = xlSheetVeryHidden
    UserForm2.show
End Sub

''''''''''''''''''''''''''''''''''''''''''UserForm2''''''''''''''''''''''''''''''''''''''''''
Private Const GWL_STYLE = -16
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000

Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare PtrSafe Function FindWindowA Lib "user32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

Public Sub SystemButtonSettings(frm As Object, show As Boolean)
Dim windowStyle As Long
Dim windowHandle As Long
windowHandle = FindWindowA(vbNullString, frm.Caption)
windowStyle = GetWindowLong(windowHandle, GWL_STYLE)
If show = False Then
    SetWindowLong windowHandle, GWL_STYLE, (windowStyle And Not WS_SYSMENU)
Else
    SetWindowLong windowHandle, GWL_STYLE, (windowStyle + WS_SYSMENU)
End If
DrawMenuBar (windowHandle)
End Sub

Private Sub CommandButton1_Click()

Dim lista As String
Dim ws As Worksheet

Set ws = Sheets("Acesso")

lista = UCase(TextBox1) & UCase(TextBox2)
On Error Resume Next
If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
ws.ShowAllData
ws.Range("A:A").AutoFilter Field:=1, Criteria1:=lista

lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row

If ws.Cells(lin_inicio, 1).Value = 0 Then
    MsgBox "USUÁRIO NÃO ENCONTRADO !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
    ws.ShowAllData
    GoTo FIM
Else
    UserForm1.show
End If

FIM:
End Sub

Private Sub CommandButton2_Click()
    ThisWorkbook.Close SaveChanges:=False
End Sub

Private Sub UserForm_Initialize()
Call SystemButtonSettings(Me, False)
End Sub

''''''''''''''''''''''''''''''''''''''''''UserForm1''''''''''''''''''''''''''''''''''''''''''
'Option Explicit
Dim gConexao As New ADODB.Connection
Dim lrs As New ADODB.Recordset
Dim strConexao, sql As String
Dim ws, wsB, wsC As Worksheet
Dim wb As Workbook

Private Sub lsConectar()
    Set gConexao = New ADODB.Connection
    
    strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\Grupos\CZ1 - Transferencia Informacoes\10. Métodos e Processos\BD_Fab_Atestado\Database_FAB_Atestado.accdb;Persist Security Info=False"
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

Private Sub CommandButton1_Click()
    Dim MyValue As Variant
    MyValue = InputBox("Digite a senha")
    If MyValue = "1010" Then
        Sheets("BD").Visible = xlSheetVisible
        Sheets("Acesso").Visible = xlSheetVisible
    Else
        MsgBox ("Senha Incorreta")
    End If
End Sub

Public Sub UserForm_Initialize()
    Sheets("BD").Visible = xlSheetVeryHidden
    Sheets("Acesso").Visible = xlSheetVeryHidden
    Unload UserForm2
    TextBox3 = Sheets("Acesso").Cells(Sheets("Acesso").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row, 2)
End Sub

Private Sub CommandButton2_Click()
    ThisWorkbook.Save
End Sub

Private Sub CommandButton3_Click()
'OK CADASTRO'
    Set ws = Sheets("BD")
    If TextBox1 = "" Or TextBox3 = "" Then
        MsgBox "FAVOR PREENCHER AS INFORMAÇÕES!", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    If OptionButton1.Value = True Then
        lsConectar
        Set lrs = New ADODB.Recordset
        lrs.Open "SELECT * FROM DB WHERE ID = " & CInt(TextBox0), gConexao, adOpenKeyset, adLockPessimistic
        
        lrs!nome = TextBox1
        lrs!setor = TextBox2
        lrs!turno = TextBox3
        lrs!DATA = TextBox4
        lrs!qtd_dias = TextBox5
        lrs!motivo = TextBox6
        lrs!obs = TextBox7
        lrs!USUARIO = Application.UserName
        lrs!edicao = Now
        
        lrs.Update
        lrs.Close
        Set lrs = Nothing
        lsDesconectar
        
        MsgBox "CADASTRO ATUALIZADO COM SUCESSO!", vbInformation
        ThisWorkbook.Save
        CommandButton4_Click
        
        UserForm1.MultiPage1.Value = 1
    Else
        lsConectar
        Set lrs = New ADODB.Recordset
        
        sql = " INSERT INTO DB "
        sql = sql & " (nome, setor, turno, data, qtd_dias, motivo, obs, usuario) "
        sql = sql & " VALUES "
        sql = sql & " ('" & TextBox1 & "', "
        sql = sql & " '" & TextBox2 & "', "
        sql = sql & " '" & TextBox3 & "', "
        sql = sql & " '" & TextBox4 & "', "
        sql = sql & " '" & TextBox5 & "', "
        sql = sql & " '" & TextBox6 & "', "
        sql = sql & " '" & TextBox7 & "', "
        sql = sql & " '" & Application.UserName & "') "
        
        lrs.Open sql, gConexao
        Set lrs = Nothing
        lsDesconectar
        
        MsgBox "CADASTRO REALIZADO COM SUCESSO!", vbInformation
        ThisWorkbook.Save
        CommandButton4_Click
        
        TextBox1.SetFocus
    End If
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub CommandButton4_Click()
    TextBox0 = ""
    TextBox1 = ""
    TextBox2 = ""
    TextBox4 = ""
    TextBox5 = ""
    TextBox6 = ""
    TextBox7 = ""
    OptionButton1.Value = False
    CheckBox1 = False
End Sub

Private Sub CommandButton5_Click()
'LIMPAR
    TextBox101 = ""
    TextBox102 = ""
    TextBox103 = ""
    TextBox104 = ""
    TextBox105 = ""
    ListBox1.Clear
    CheckBox2 = False
End Sub

Private Sub CommandButton6_Click()
'OK PESQUISA'
    Application.ScreenUpdating = False
    Set ws = Sheets("BD")
    
    If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    ws.ShowAllData
    ws.Rows("2:" & ws.Cells(1, 1).End(xlDown).Row).ClearContents
    
    lsConectar
    Set lrs = New ADODB.Recordset
    lrs.Open " SELECT * FROM DB ", gConexao, adOpenKeyset, adLockPessimistic
    ws.Cells(2, 1).CopyFromRecordset lrs
    ws.Columns("A:A").NumberFormat = "0"
    ws.Columns("E:E").NumberFormat = "dd/mm/yyyy"
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
    
    Set rngAF = ws.Range("A1:A" & ws.Cells(1, 1).End(xlDown).Row)
    
    ws.Range("D:D").AutoFilter Field:=4, Criteria1:=TextBox3
    If TextBox101 <> "" Then ws.Range("B:B").AutoFilter Field:=2, Criteria1:="=*" & TextBox101 & "*"
    If TextBox102 <> "" Then ws.Range("C:C").AutoFilter Field:=3, Criteria1:="=*" & TextBox102 & "*"
    If TextBox103 <> "" And TextBox104 <> "" Then ws.Range("E:E").AutoFilter Field:=5, Criteria1:=">=" & Format(CDate(TextBox103), "mm/dd/yyyy"), Operator:=xlAnd, Criteria2:="<=" & Format(CDate(TextBox104), "mm/dd/yyyy")
    If TextBox105 <> "" Then ws.Range("G:G").AutoFilter Field:=7, Criteria1:="=*" & TextBox105 & "*"
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    If ws.Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        ListBox1.Clear
        ws.ShowAllData
        GoTo FIM
    End If

    Dim arrayItems2()
    With Planilha2
        ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, ws.Range("A:A")), 1 To 8)
        Me.ListBox1.ColumnCount = 8
        Me.ListBox1.ColumnWidths = "40;300;;;;;;;"
        i = 0
        For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
            Me.ListBox1.AddItem
            For coluna = 1 To 8
                arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
            Next coluna
            i = i + 1
        Next rngcell
        Me.ListBox1.List = arrayItems2()
    End With
    
FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton7_Click()
'EDITAR PESQUISA
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR REALIZAR UMA PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Set ws = Sheets("BD")
    ID = ListBox1.List(ListBox1.ListIndex, 0)
    If ID = 0 Or Not IsNumeric(ID) Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "SELECIONE UM ITEM VÁLIDO PARA EDITAR", vbCritical
        GoTo FIM
    End If
    If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    ws.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    TextBox0 = ID
    TextBox1 = ws.Cells(lin_inicio, 2)
    TextBox2 = ws.Cells(lin_inicio, 3)
    TextBox3 = ws.Cells(lin_inicio, 4)
    TextBox4 = ws.Cells(lin_inicio, 5)
    TextBox5 = ws.Cells(lin_inicio, 6)
    TextBox6 = ws.Cells(lin_inicio, 7)
    TextBox7 = ws.Cells(lin_inicio, 8)
    OptionButton1.Value = True
    
    ws.ShowAllData
    
    ListBox1.Clear
    UserForm1.MultiPage1.Value = 0
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub CommandButton8_Click()
'DELETAR
    Application.ScreenUpdating = False
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR REALIZAR UMA PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Set ws = Sheets("BD")
    ID = ListBox1.List(ListBox1.ListIndex, 0)
    
    result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR UM REGISTRO?", vbYesNo + vbCritical)
    If result = vbYes Then
        ws.Range("A:A").AutoFilter Field:=1, Criteria1:=ID
        lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        lsConectar
        Set lrs = New ADODB.Recordset
        lrs.Open "SELECT * FROM DB WHERE ID = " & CInt(ID), gConexao, adOpenKeyset, adLockPessimistic
        lrs.Delete
        lrs.Update
        Set lrs = Nothing
        
        Set lrs = New ADODB.Recordset
        sql = " INSERT INTO Excluir "
        sql = sql & " (nome, setor, turno, data, qtd_dias, motivo, obs, usuario, edicao, excluido) "
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
        sql = sql & " '" & Now & "') "
        
        lrs.Open sql, gConexao
'        lrs.Close
        Set lrs = Nothing
        lsDesconectar
        
        MsgBox "CADASTRO EXCLUIDO COM SUCESSO!", vbInformation
        ListBox1.Clear
    End If
FIM:
Application.ScreenUpdating = True
End Sub

Private Sub CheckBox1_Click()
    If CheckBox1 = True Then
        TextBox4 = Date
    Else
        TextBox4 = ""
    End If
End Sub

Private Sub CheckBox2_Click()
    If CheckBox2 = True Then
        TextBox104 = Date
    Else
        TextBox104 = ""
    End If
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextBox4.MaxLength = 10
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        If TextBox4.SelStart = 2 Then
            TextBox4.SelText = "/"
        End If
        If TextBox4.SelStart = 5 Then
            TextBox4.SelText = "/"
        End If
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox6_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox7_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox101_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox102_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox105_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub TextBox103_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextBox103.MaxLength = 10
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        If TextBox103.SelStart = 2 Then
            TextBox103.SelText = "/"
        End If
        If TextBox103.SelStart = 5 Then
            TextBox103.SelText = "/"
        End If
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub
Private Sub TextBox104_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextBox104.MaxLength = 10
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        If TextBox104.SelStart = 2 Then
            TextBox104.SelText = "/"
        End If
        If TextBox104.SelStart = 5 Then
            TextBox104.SelText = "/"
        End If
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub
