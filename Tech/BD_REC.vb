---------------------------------------------------MÓDULO1----------------------------------------------------
Sub inserir()
Load UserForm1
With UserForm1
    .Top = Application.Top + 40
    .Left = Application.Left + 950
    .Height = 225
    .Width = 305
End With
UserForm1.Show
End Sub
----------------------------------------------------------------------------------------------------------------
Public Sub UserForm_Initialize()
    TextBox4 = Format(CDate(Now()), "dd/mm/yyyy")
End Sub
Private Sub CommandButton1_Click()
Application.ScreenUpdating = False
Application.DisplayStatusBar = False
Application.EnableEvents = False
    If TextBox3 = "" Then
        MsgBox "Favor preencher as informações para continuar", vbExclamation
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    Set ws = Sheets(1)
    ws.Range("B5").Select
    ws.ListObjects(1).ShowAutoFilter = True
    ws.ListObjects(1).AutoFilter.ShowAllData
'    If TextBox1 <> "" Then ws.ListObjects(1).Range.AutoFilter Field:=4, Criteria1:=TextBox1
'    If TextBox2 <> "" Then ws.ListObjects(1).Range.AutoFilter Field:=5, Criteria1:=TextBox2
    ws.ListObjects(1).Range.AutoFilter Field:=8, Criteria1:=TextBox3
    If ws.Cells(ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row, 2).Value = 0 Then
        ws.ListObjects(1).AutoFilter.ShowAllData
        lin = ws.Cells(5, 2).End(xlDown).Row + 1
        ws.Cells(lin, 2) = Application.WorksheetFunction.Max(ws.Range("B:B")) + 1
        ws.Cells(lin, 3) = 0
    Else
        rec = ws.Cells(ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row, 2).Value
        ws.ListObjects(1).AutoFilter.ShowAllData
        ws.ListObjects(1).Range.AutoFilter Field:=1, Criteria1:=rec
        rev = Application.WorksheetFunction.Max(ws.Range("C:C").SpecialCells(xlCellTypeVisible)) + 1
        ws.ListObjects(1).AutoFilter.ShowAllData
        lin = ws.Cells(5, 2).End(xlDown).Row + 1
        ws.Cells(lin, 2) = rec
        ws.Cells(lin, 3) = rev
    End If
    ws.Cells(lin, 4) = TextBox4
    ws.Cells(lin, 5) = TextBox1
    ws.Cells(lin, 6) = TextBox2
    ws.Cells(lin, 7) = Left(Right(TextBox3, 12), 9)
    ws.Cells(lin, 9) = TextBox3
    ws.Cells(lin, 13) = Application.UserName
    MsgBox "Cadastro realizado com sucesso", vbInformation
    
FIM:
Application.ScreenUpdating = True
Application.DisplayStatusBar = True
Application.EnableEvents = True
End Sub
Private Sub CommandButton2_Click()
    Unload UserForm1
End Sub
Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextBox1.MaxLength = 5
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub
Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    TextBox2.MaxLength = 6
End Sub
Private Sub TextBox2_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not TextBox2 = "" Then
        TextBox2 = UCase(Replace(Replace(Replace(TextBox2, " ", vbNullString), ".", vbNullString), "-", vbNullString))
        If Not InStr(1, TextBox2, "E") > 0 And IsNumeric(Left(TextBox2, 1)) Then
            TextBox2 = "E-" & Left(TextBox2, 1) & "." & Right(Left(TextBox2, 3), 2)
        Else
            TextBox2 = Left(TextBox2, 1) & "-" & Right(Left(TextBox2, 2), 1) & "." & Right(Left(TextBox2, 4), 2)
        End If
    End If
End Sub
Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    TextBox3.MaxLength = 18
End Sub
Private Sub TextBox3_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Not TextBox3 = "" Then
        TextBox3 = UCase(Replace(Replace(TextBox3, " ", vbNullString), ".", vbNullString))
        If Not InStr(1, TextBox3, "BM") > 0 And IsNumeric(Left(TextBox3, 1)) Then
            If Len(TextBox3) < 11 Then
                TextBox3 = "BM " & Left(TextBox3, 3) & "." & Right(Left(TextBox3, 8), 5) & "." & Right(Left(TextBox3, 10), 2)
            Else
                TextBox3 = "BM " & Left(TextBox3, 3) & "." & Right(Left(TextBox3, 8), 5) & "." & Right(Left(TextBox3, 10), 2) & " " & Right(TextBox3, 2)
            End If
        Else
            If Len(TextBox3) < 13 Then
                TextBox3 = Left(TextBox3, 2) & " " & Right(Left(TextBox3, 5), 3) & "." & Right(Left(TextBox3, 10), 5) & "." & Right(Left(TextBox3, 12), 2)
            Else
                TextBox3 = Left(TextBox3, 2) & " " & Right(Left(TextBox3, 5), 3) & "." & Right(Left(TextBox3, 10), 5) & "." & Right(Left(TextBox3, 12), 2) & " " & Right(TextBox3, 2)
            End If
        End If
    End If
End Sub
'Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
'    TextBox3.MaxLength = 18
'    Select Case KeyAscii
'        Case 8 'Aceita o BACK SPACE
'        Case 13: SendKeys "{TAB}" 'Emula o TAB
'        Case 48 To 57, 65 To 90
'        If TextBox3.SelStart = 2 Then TextBox3.SelText = " "
'        If TextBox3.SelStart = 6 Then TextBox3.SelText = "."
'        If TextBox3.SelStart = 12 Then TextBox3.SelText = "."
'        If TextBox3.SelStart = 15 Then TextBox3.SelText = " "
'    Case Else: KeyAscii = 0 'Ignora os outros caracteres
'    End Select
'End Sub
'Private Sub TextBox3_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
'    If KeyCode = 86 And Shift = 2 Then 'BLOQUEIA O CRTL+V
'        KeyCode = 0
'    End If
'End Sub
