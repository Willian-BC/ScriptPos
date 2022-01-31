Public Sub UserForm_Initialize()
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    Sheets("Baixa").Visible = xlSheetVeryHidden
    Sheets("Excluido").Visible = xlSheetVeryHidden
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
