Private Sub TextBox103_Exit(ByVal Cancel As MSForms.ReturnBoolean)
'IMPORTAÇÃO OP CADASTRO
    If TextBox109 = "" Then
        Application.ScreenUpdating = False
        Sheets("Componentes").Visible = xlSheetVisible
        Sheets("Componentes").Activate
        
        On Error Resume Next
        Worksheets("Componentes").ListObjects(1).ShowAutoFilter = True
        Worksheets("Componentes").ListObjects(1).AutoFilter.ShowAllData
        
        Worksheets("Componentes").ListObjects(1).Range.AutoFilter Field:=5, Criteria1:=TextBox103.Text
        
        lin_inicio = Sheets("Componentes").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        If Cells(lin_inicio, 1).Value = 0 Then
            Worksheets("Componentes").ListObjects(1).AutoFilter.ShowAllData
            TextBox101 = ""
            TextBox111 = ""
            TextBox102 = ""
            TextBox104 = ""
            TextBox105 = ""
            GoTo FIM
        End If
        
        TextBox101 = Sheets("Componentes").Cells(lin_inicio, 2)
        TextBox111 = Sheets("Componentes").Cells(lin_inicio, 3)
        TextBox102 = Sheets("Componentes").Cells(lin_inicio, 4)
        TextBox104 = Sheets("Componentes").Cells(lin_inicio, 6)
        TextBox105 = Sheets("Componentes").Cells(lin_inicio, 7)
        
        Worksheets("Componentes").ListObjects(1).AutoFilter.ShowAllData
        
FIM:
        Sheets("Componentes").Visible = xlSheetVeryHidden
        TextBox106.SetFocus
        Application.ScreenUpdating = True
    End If
End Sub
