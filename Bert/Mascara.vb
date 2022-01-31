Private Sub TextBox204_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
  
''''''''''''''''''LETRA MAIUSCULA'''''''''''''''''''''''''''
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
  
''''''''''''''''''1º CARACTER NUM'''''''''''''''''''''''''''
    If IsNumeric(Left(TextBox204, 1)) Then
    
''''''''''''''''''COMPRIMENTO'''''''''''''''''''''''''''
        TextBox204.MaxLength = 8
        Select Case KeyAscii
            Case 8 'Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}" 'Emula o TAB
            Case 48 To 57
            If TextBox204.SelStart = 2 Then
                TextBox204.SelText = "-"
            End If
            If TextBox204.SelStart = 5 Then
                TextBox204.SelText = "-"
            End If
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
        End Select
    Else
        TextBox204.MaxLength = 7
        Select Case KeyAscii
            Case 8 'Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}" 'Emula o TAB
      
''''''''''''''''''RESTRIÇÃO CARACTERES'''''''''''''''''''''''''''
    Case 48 To 57 , 65 To 90
            If TextBox204.SelStart = 1 Then
                TextBox204.SelText = "-"
            End If
            If TextBox204.SelStart = 4 Then
                TextBox204.SelText = "-"
            End If
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
        End Select
    End If
End Sub
