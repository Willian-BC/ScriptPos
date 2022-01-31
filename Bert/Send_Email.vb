Sub ENVIAR()
    
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    ThisWorkbook.Save
    
    With OutlookMail
                .To = "ticket@outlook.com.br"
        .Subject = "Ticket"
        .Body = "Prezados," & vbNewLine & "Segue ticket para avaliação." & vbNewLine _
            & vbNewLine & "Nome: " & Sheets("Planilha1").Range("B3").Value _
            & vbNewLine & "Setor: " & Sheets("Planilha1").Range("D3").Value _
            & vbNewLine & "Descrição: " & Sheets("Planilha1").Range("B5").Value _
            & vbNewLine & vbNewLine & vbNewLine & "Att,"
        .Attachments.Add ActiveWorkbook.FullName
        .Send
    End With
    
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    
End Sub
