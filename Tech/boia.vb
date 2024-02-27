Sub inserir()
Load UserForm1
With UserForm1
    .Top = Application.Top + 40
    .Left = Application.Left + 950
    .Height = 335
    .Width = 480
End With
UserForm1.Show
End Sub
-------------------------------------------------------------------------------------------------
Public Sub UserForm_Initialize()
    Set ws = Sheets("MARCAÇÃO")
    TextBox1 = ws.Cells(2, 4).Value
    TextBox2 = ws.Cells(4, 6).Value
    Application.DecimalSeparator = ","
    TextBox3 = ws.Cells(16, 14).Value
    TextBox4 = ws.Cells(17, 14).Value
    TextBox5 = ws.Cells(18, 14).Value
    TextBox6 = ws.Cells(19, 14).Value
    TextBox7 = ws.Cells(20, 14).Value
    TextBox8 = ws.Cells(21, 14).Value
End Sub
  
Private Sub CommandButton1_Click()
    Application.ScreenUpdating = False
    If TextBox1 = "" Or TextBox2 = "" Or TextBox3 = "" Or TextBox4 = "" Or TextBox5 = "" Or TextBox6 = "" Or TextBox7 = "" Or TextBox8 = "" Then
        MsgBox "Favor preencher todas as informações para continuar", vbExclamation
        GoTo FIM
    End If
    Set ws = Sheets("MARCAÇÃO")
    ws.Range("B15:F50").ClearContents
    ws.Range("B56:F96").ClearContents
    ws.Rows("98:" & ws.Cells(1, 1).End(xlDown).Row).Delete
    
    ws.Cells(2, 4) = TextBox1
    ws.Cells(4, 6) = TextBox2
    ws.Cells(12, 6) = TextBox3
    ws.Cells(16, 14) = TextBox3
    ws.Cells(17, 14) = TextBox4
    ws.Cells(8, 6) = TextBox5
    ws.Cells(18, 14) = TextBox5
    ws.Cells(10, 6) = TextBox6
    ws.Cells(19, 14) = TextBox6
    ws.Cells(20, 14) = TextBox7
    ws.Cells(21, 14) = TextBox8
    
    If CInt(TextBox8) > 18 Then
        fol = WorksheetFunction.RoundUp((CInt(TextBox8) - 18) / 20, 0) + 1
        If fol > 2 Then
            j = 98
            For i = 3 To fol
                Set rg = ws.Range("B52:I97")
                rg.Copy
                ws.Range("B" & j).PasteSpecial Paste:=xlPasteAll
                j = j + 46
                ws.PageSetup.PrintArea = "$B$1:$I$" & j
                Set ws.HPageBreaks(i - 1).Location = Range("B" & j - 46)
            Next i
        End If
    Else
        fol = 1
    End If
    lin = 16
    k = 50
    For i = 1 To CInt(TextBox8)
        If i = 1 Then
            ws.Cells(lin, 2) = CInt(i)
            ws.Cells(lin, 4) = (CDbl(TextBox3) + CDbl(TextBox4)) - CDbl(TextBox6)
            ws.Cells(lin - 1, 5) = (WorksheetFunction.Quotient(ws.Cells(lin, 4), 10) * 10) + 10
            ws.Cells(lin + 1, 3) = CDbl(TextBox7)
        Else
            If lin = k Then
                lin = lin + 6
                ws.Cells(lin, 2) = CInt(i)
                ws.Cells(lin, 4) = ws.Cells(lin - 8, 4) - CDbl(TextBox7)
                ws.Cells(lin + 1, 3) = CDbl(TextBox7)
                If (ws.Cells(lin, 4) < (WorksheetFunction.Quotient(ws.Cells(lin, 4), 10) * 10) + 10) And ((WorksheetFunction.Quotient(ws.Cells(lin, 4), 10) * 10) + 10 < ws.Cells(lin - 8, 4)) Then
                    ws.Cells(lin - 7, 5) = (WorksheetFunction.Quotient(ws.Cells(lin, 4), 10) * 10) + 10
                End If
                k = k + 46
            Else
                ws.Cells(lin, 2) = CInt(i)
                ws.Cells(lin, 4) = ws.Cells(lin - 2, 4) - CDbl(TextBox7)
                ws.Cells(lin + 1, 3) = CDbl(TextBox7)
                If (ws.Cells(lin, 4) < (WorksheetFunction.Quotient(ws.Cells(lin, 4), 10) * 10) + 10) And ((WorksheetFunction.Quotient(ws.Cells(lin, 4), 10) * 10) + 10 < ws.Cells(lin - 2, 4)) Then
                    ws.Cells(lin - 1, 5) = (WorksheetFunction.Quotient(ws.Cells(lin, 4), 10) * 10) + 10
                End If
            End If
        End If
        lin = lin + 2
        
    Next i
    ws.Cells(lin - 1, 5) = WorksheetFunction.Quotient(ws.Cells(lin - 2, 4), 10) * 10
FIM:
    Application.ScreenUpdating = True
    If CheckBox1 = True Then Unload Me
End Sub

Private Sub CommandButton2_Click()
    TextBox1 = ""
    TextBox2 = ""
    TextBox3 = ""
    TextBox4 = ""
    TextBox5 = ""
    TextBox6 = ""
    TextBox7 = ""
    TextBox8 = ""
End Sub

Private Sub CommandButton3_Click()
    Unload Me
End Sub

