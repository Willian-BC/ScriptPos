Sub Consultar()

    Dim IE As Object
    Set IE = CreateObject("InternetExplorer.Application")
    Dim doc As HTMLDocument
    Dim allRowofData As Object
    Dim W As Worksheet
    Dim Ultcel As Range
    Dim col As Integer
    Dim ln As Long
    Dim situacao As String
    Dim aviso
       
       
    Set W = Planilha1
    W.Range("A2").Select
    
    Set Ultcel = W.Cells(W.Rows.Count, 1).End(xlUp)
        
        
   
    'Abre internet explorer
    IE.Visible = False
    IE.navigate "https://www.sefaz.go.gov.br/ccn/"
    
    'Espera segundos antes de iniciar busca
    Do While IE.Busy
        Application.Wait DateAdd("s", 1, Now)
    Loop
    
    If Cells(2, 2).Value <> "" Then
        ln = Cells(2, 2).End(xlDown).Row + 1
    Else
        ln = 2
    End If
    col = 1
    
    'Insere e copia dados
    
    Do While ln <= Ultcel.Row
 
    
    
    Set doc = IE.document
    
    doc.getElementsByName("tipoDocumento")(1).Checked = True
    doc.getElementById("numrDocumento").Value = W.Cells(ln, col)
    doc.getElementById("btnSubmit").Click
    
    Do While IE.Busy
        Application.Wait DateAdd("s", 3, Now)
    Loop
    
        situacao = doc.frames(0).document.getElementsByTagName("td")(1).innerText
        
        W.Cells(ln, col + 1) = situacao
        
    ln = ln + 1
    
    On Error GoTo aviso
    
    Loop
    
     
    
    IE.Quit
    
  '  W.UsedRange.EntireColumn.AutoFit
    
    Set IE = Nothing
    
   Exit Sub
   
aviso:            MsgBox "CNPJ não cadastrado ou inválido, delete o valor para continuar"
     
End Sub
