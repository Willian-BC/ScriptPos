Function NomeUsuario()

NomeUsuario = Application.UserName

End Function
Private Sub FAR()
Application.ScreenUpdating = False
    Dim FindBD As String, ReplaceBD As String
    Dim i As Byte
    Dim cht As ChartObject
    Dim ser As Series
    
    i = ThisWorkbook.Sheets.Count
    Sheets("BD TEMPLATE").Copy After:=Sheets(i - 2)
    ActiveSheet.Name = "TEMPLATE BDX"
    Sheets("TEMPLATE BDX").Tab.Color = xlAutomatic
    Sheets("ST TEMPLATE").Copy After:=Sheets(i - 1)
    ActiveSheet.Name = "TEMPLATE STX"
    Sheets("TEMPLATE STX").Tab.Color = xlAutomatic

    FindBD = "BD TEMPLATE"
    ReplaceBD = Sheets(i + 1).Name
    ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 19")).Delete
    ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 21")).Delete
    Columns("E:F").Replace What:=FindBD, replacement:=ReplaceBD, LookAt:=xlPart, MatchCase:=True
    ActiveSheet.Cells(3, 8) = Application.UserName
    ActiveSheet.Cells(2, 7) = Now()
    ActiveCell.NumberFormat = "dd/m/yyyy h:mm;@"

    For Each cht In ActiveSheet.ChartObjects
        For Each ser In cht.Chart.SeriesCollection
            On Error GoTo PROX1
            ser.Formula = WorksheetFunction.Substitute(ser.Formula, FindBD, ReplaceBD)
        Next ser
PROX1:
    Resume PROX2
PROX2:
    Next cht


Application.ScreenUpdating = True
End Sub
Private Sub PESQUISA()
    Dim i, a As Byte
    i = ThisWorkbook.Sheets.Count
    Sheets("LISTA DE ACESSO").Columns("J:J").Clear
    For a = 3 To i
        Sheets("LISTA DE ACESSO").Cells(a, 10) = Sheets(a).Name
    Next
    
    UserForm2.Show
End Sub

Private Sub GRAFICO()
    Dim lG2 As Byte
    
    Columns("J:N").Clear
    
    Cells(2, 10).Value = "Gaiola"
    Cells(2, 11).Value = "Ângulo Puxada"
    Cells(2, 12).Value = "Ângulo Torção"
    Cells(2, 13).Value = "Média Amplitude"
    Cells(2, 14).Value = "Média Passo"
    
lG2 = Cells(2, 6).End(xlDown).Row
lG2 = Cells(lG2, 6).End(xlDown).Row
    Cells(lG2, 10).Value = "Gaiola"
    Cells(lG2, 11).Value = "Ângulo Puxada"
    Cells(lG2, 12).Value = "Ângulo Torção"
    Cells(lG2, 13).Value = "Média Amplitude"
    Cells(lG2, 14).Value = "Média Passo"
lG2 = lG2 + 1

If Cells(3, 8).Value <> "" Then
    Range("F3", Cells(3, 6).End(xlDown)).Copy
    Range("J3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    ActiveSheet.Range("J2:J" & Cells(3, 10).End(xlDown).Row).RemoveDuplicates Columns:=1, Header:=xlYes
    Range("E3", Cells(3, 5).End(xlDown)).Copy
    Range("K3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    ActiveSheet.Range("K2:K" & Cells(3, 11).End(xlDown).Row).RemoveDuplicates Columns:=1, Header:=xlYes
    Range("G3", Cells(3, 7).End(xlDown)).Copy
     Range("L3").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    ActiveSheet.Range("L2:L" & Cells(3, 12).End(xlDown).Row).RemoveDuplicates Columns:=1, Header:=xlYes
    'ltorcao/lpuxada = Cells(3, 11).End(xlDown).Row     'Range("J3", Cells(3, 10).End(xlDown)).count

    puxadaG1 = Application.Count(Columns("K"))
    torcaoG1 = Application.Count(Columns("L"))
'    aG1 = Cells(3, 11).End(xlDown).Row
    If Cells(3, 12).End(xlDown).Row < 5000 Then
        bg1 = Cells(3, 12).End(xlDown).Row
    Else
        bg1 = 3
    End If
    cG1 = 4
    For aux = 2 To puxadaG1
        Cells(cG1, 11).Cut Cells(Cells(3, 12).End(xlDown).Row + 1, 11)
        Range("L3:L" & bg1).Copy
        Range("L" & Cells(3, 12).End(xlDown).Row + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=True, Transpose:=False
        cG1 = cG1 + 1
    Next
    aux3 = 3
    aux4 = 3
    For aux1 = 1 To puxadaG1
        For aux2 = 1 To torcaoG1
        On Error Resume Next
            Range("M" & aux3) = Application.WorksheetFunction.AverageIfs(Range("H:H"), Range("G:G"), _
            Range("L" & aux3).Value, Range("E:E"), Range("K" & aux4).Value, Range("F:F"), Range("J3").Value)
            Range("N" & aux3) = Application.WorksheetFunction.AverageIfs(Range("I:I"), Range("G:G"), _
            Range("L" & aux3).Value, Range("E:E"), Range("K" & aux4).Value, Range("F:F"), Range("J3").Value)
            aux3 = aux3 + 1
        Next
        aux4 = aux3
    Next
End If
If Cells(lG2, 8).Value <> "" Then
    Range("F" & lG2, Cells(lG2, 6).End(xlDown)).Copy
    Range("J" & lG2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    ActiveSheet.Range("J" & lG2 - 1 & ":J" & Cells(lG2, 10).End(xlDown).Row).RemoveDuplicates Columns:=1, Header:=xlYes
    Range("E" & lG2, Cells(lG2, 5).End(xlDown)).Copy
    Range("K" & lG2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    ActiveSheet.Range("K" & lG2 - 1 & ":K" & Cells(lG2, 11).End(xlDown).Row).RemoveDuplicates Columns:=1, Header:=xlYes
    Range("G" & lG2, Cells(lG2, 7).End(xlDown)).Copy
     Range("L" & lG2).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=True, Transpose:=False
    ActiveSheet.Range("L" & lG2 - 1 & ":L" & Cells(lG2, 12).End(xlDown).Row).RemoveDuplicates Columns:=1, Header:=xlYes
    'ltorcao/lpuxada = Cells(3, 11).End(xlDown).Row     'Range("J3", Cells(3, 10).End(xlDown)).count

    puxadaG2 = Application.Count(Columns("K")) - puxadaG1
    torcaoG2 = Application.Count(Columns("L")) - (torcaoG1 * puxadaG1)
'    aG2 = Cells(lG2, 11).End(xlDown).Row
    If Cells(lG2, 12).End(xlDown).Row < 8000 Then
        bg2 = Cells(lG2, 12).End(xlDown).Row
    Else
        bg2 = lG2
    End If
    cG2 = lG2 + 1
    For aux = 2 To puxadaG1
        Cells(cG2, 11).Cut Cells(Cells(lG2, 12).End(xlDown).Row + 1, 11)
        Range("L" & lG2 & ":L" & bg2).Copy
        Range("L" & Cells(lG2, 12).End(xlDown).Row + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=True, Transpose:=False
        cG2 = cG2 + 1
    Next
    aux7 = lG2
    aux8 = lG2
    For aux5 = 1 To puxadaG2
        For aux6 = 1 To torcaoG2
        On Error Resume Next
            Range("M" & aux7) = Application.WorksheetFunction.AverageIfs(Range("H:H"), Range("G:G"), _
            Range("L" & aux7).Value, Range("E:E"), Range("K" & aux8).Value, Range("F:F"), Range("J" & lG2).Value)
            Range("N" & aux7) = Application.WorksheetFunction.AverageIfs(Range("I:I"), Range("G:G"), _
            Range("L" & aux7).Value, Range("E:E"), Range("K" & aux8).Value, Range("F:F"), Range("J" & lG2).Value)
            aux7 = aux7 + 1
        Next
        aux8 = aux7
    Next
End If
'        If Range("J3", Cells(3, 10).End(xlDown)).Count > 1 Then
'            If Range("J3", Cells(3, 10).End(xlDown)).Count > 2 Then
'                Range("J4", Cells(3, 10).End(xlDown)).Cut
'                Cells((Cells(3, 11).End(xlDown).Row) + 1, 10).Select
'                ActiveSheet.Paste
'                lpuxada = Cells((Cells(3, 11).End(xlDown).Row) + 1, 10).End(xlDown).Row
'                Range("K3:K" & ltorcao).Copy
'                Range("K" & Cells(3, 11).End(xlDown).Row + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                :=True, Transpose:=False
'                GoTo fim
'            End If
'            Range("J4").Cut
'            Cells(ltorcao + 1, 10).Select
'            ActiveSheet.Paste
'            Range("K3:K" & linha).Copy
'            Range("K" & Cells(3, 11).End(xlDown).Row + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'            :=True, Transpose:=False
'        End If
'fim:
'        a = ltorcao + 1
'        ltorcao = Cells(3, 11).End(xlDown).Row
        
'        Range("J" & Cells(3, 10).End(xlDown).Row).Cut
'        Cells((Cells(3, 11).End(xlDown).Row) + 1, 10).Select
'        ActiveSheet.Paste
'        Range("K3:K" & linha).Copy
'        Range("K" & Cells(3, 11).End(xlDown).Row + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'        :=True, Transpose:=False
'    Next
    
End Sub

'Codigo Userfor1 (login e senha)

Private Declare Function FindWindowA Lib "USER32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLongA Lib "USER32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "USER32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Sub cmdacessar_Click()

Dim DATA As Date
Dim USUARIO, SENHA As String

USUARIO = Textusuario.Text
SENHA = Textsenha.Text

LISTAUSUARIO = USUARIO & " " & SENHA
ThisWorkbook.Activate

On Error Resume Next
With Sheets("LISTA DE ACESSO")
    Set PESQ = .Cells.Find(What:=LISTAUSUARIO, After:=.Cells(7), LookIn:=xlValues, LookAt:= _
        xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False)
End With

nao = 0
If USUARIO = "WCARDOSO" And SENHA = "1234" Then
    Load Me
    Sheets("LISTA DE ACESSO").Visible = xlSheetVisible     'EXIBINDO A ABA CONTROLE DE ACESSO
    ThisWorkbook.Activate
    Unload UserForm1
    i = (Sheets("LISTA DE ACESSO").Cells(Rows.Count, 1).End(xlUp).Row) + 1
    Sheets("LISTA DE ACESSO").Cells(i, 1) = Textusuario.Text
    Sheets("LISTA DE ACESSO").Cells(i, 2) = NomeUsuario()
    Sheets("LISTA DE ACESSO").Cells(i, 3) = TextBoxdata.Text
    TextBoxdata = Format("dd/m/yyyy h:mm;@")
    Application.Visible = True
ElseIf PESQ Is Nothing Then
    MsgBox "USUÁRIO OU SENHA INCORRETO", vbExclamation, "ATENÇÃO!"
    Textusuario.Text = ""
    Textsenha.Text = ""
    nao = 1
Else
    Load Me
    ThisWorkbook.Activate
    Unload UserForm1
    i = (Sheets("LISTA DE ACESSO").Cells(Rows.Count, 1).End(xlUp).Row) + 1
    Sheets("LISTA DE ACESSO").Cells(i, 1) = Textusuario.Text
    Sheets("LISTA DE ACESSO").Cells(i, 2) = NomeUsuario()
    Sheets("LISTA DE ACESSO").Cells(i, 3) = TextBoxdata.Text
    TextBoxdata = Format("dd/m/yyyy h:mm;@")
    Application.Visible = True
End If

If nao = 0 Then
    i = ThisWorkbook.Sheets.Count
    For a = 1 To i
        If Sheets(a).Name = "START" Then
            Sheets("BD TEMPLATE").Visible = xlSheetVisible
            Sheets("ST TEMPLATE").Visible = xlSheetVisible
            Sheets(a).Visible = xlSheetVeryHidden
        ElseIf Sheets(a).Name <> "BD TEMPLATE" And Sheets(a).Name <> "ST TEMPLATE" And Sheets(a).Name <> "LISTA DE ACESSO" Then
            Sheets(a).Visible = xlHidden
        End If
    Next
End If
nao = 0
'ThisWorkbook.Save  salvar quem acessou a planilha
'UserForm2.Show
End Sub
Private Sub cmdcancelar_Click()
Unload Me
Application.Visible = True
ThisWorkbook.Close SaveChanges:=False
End Sub
Private Sub Textusuario_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Textusuario.Text = UCase(Textusuario.Text)
End Sub
Private Sub UserForm_Initialize()
'    i = ThisWorkbook.Sheets.Count
'    SheetSheets("LISTA DE ACESSO").Visible = xlSheetVeryHidden
'    For a = 2 To i - 2
'        Sheets(a).Visible = xlHidden
'    Next
    TextBoxdata = Now()
    Dim hwnd As Long
    hwnd = FindWindowA(vbNullString, Me.Caption)
    SetWindowLongA hwnd, -16, GetWindowLongA(hwnd, -16) And &HFFF7FFFF
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim hwnd As Long
    hwnd = FindWindowA(vbNullString, Me.Caption)
    SetWindowLongA hwnd, -16, GetWindowLongA(hwnd, -16) Or &H80000
End Sub

'Codigo Userfor2 (pesquisa)

Private Declare Function FindWindowA Lib "USER32" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLongA Lib "USER32" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLongA Lib "USER32" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Sub CommandButton1_Click()
    Dim ABA As String
    
    If IsNull(ListBox1.Object) Then
        MsgBox "SELECIONE UM ÍTEM", vbCritical, "ATENÇÃO!"
    Else
    ABA = ListBox1.Object
    Sheets(ABA).Visible = xlSheetVisible
    Sheets(ABA).Select
    Unload Me
    End If
End Sub

Private Sub CommandButton2_Click()
    Unload Me
End Sub

Private Sub TextBox1_Change()
    sValues = Application.Transpose(Sheets("LISTA DE ACESSO").Range("J3:J" & Sheets("LISTA DE ACESSO").Range("J1").End(xlDown).Row).Value)
    ListBox1.List = Filter(sValues, TextBox1.Text, True, vbTextCompare)
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    Dim hwnd As Long
    hwnd = FindWindowA(vbNullString, Me.Caption)
    SetWindowLongA hwnd, -16, GetWindowLongA(hwnd, -16) Or &H80000
End Sub

Private Sub UserForm_Initialize()
    Dim sValues() As Variant
    sValues = Application.Transpose(Sheets("LISTA DE ACESSO").Range("J3:J" & Sheets("LISTA DE ACESSO").Range("J3").End(xlDown).Row).Value)
    ListBox1.List = sValues
    
End Sub
