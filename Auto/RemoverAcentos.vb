Function fnRetirarAcentos(ByVal vStrPalavra As String) As String
    Dim lstrEspecial    As String
    Dim lstrSubstituto  As String
    Dim lstrAlterada    As String
    Dim liControle      As Integer
    Dim liPosicao       As Integer
    Dim lstrLetra       As String
    
    Application.Volatile

    lstrEspecial = "àáâãäèéêëìíîïòóôõöùúûüÀÁÂÃÄÈÉÊËÌÍÎÒÓÔÕÖÙÚÛÜçÇñÑ"
    lstrSubstituto = "aaaaaeeeeiiiiooooouuuuAAAAAEEEEIIIOOOOOUUUUcCnN"
 
    lstrAlterada = ""
 
    If vStrPalavra <> "" Then
        For liControle = 1 To Len(vStrPalavra)
            lstrLetra = Mid(vStrPalavra, liControle, 1)
            liPosicao = InStr(lstrEspecial, lstrLetra)
        
            If liPosicao > 0 Then
                lstrLetra = Mid(lstrSubstituto, liPosicao, 1)
            End If
        
            lstrAlterada = lstrAlterada & lstrLetra
        Next
        
        fnRetirarAcentos = lstrAlterada
    End If
End Function
