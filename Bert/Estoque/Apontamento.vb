Private Sub CommandButton301_Click()
'OK IMPORTAÇÃO
    If TextBox301 = "" Or TextBox302 = "" Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS!", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Sheets("Componentes").Visible = xlSheetVisible
    Sheets("Componentes").Activate
    
    On Error Resume Next
    Worksheets("Componentes").ListObjects(1).ShowAutoFilter = True
    Worksheets("Componentes").ListObjects(1).AutoFilter.ShowAllData
    Set rngAF = Range("A1:A" & Cells(1, 1).End(xlDown).Row)
    
    Worksheets("Componentes").ListObjects(1).Range.AutoFilter Field:=4, Criteria1:=TextBox301.Text
    Worksheets("Componentes").ListObjects(1).Range.AutoFilter Field:=3, Criteria1:="=*" & TextBox302.Text & "*"
    
    lin_inicio = Sheets("Componentes").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = Sheets("Componentes").Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If Sheets("Componentes").Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        Sheets("Base de dados").ShowAllData
        GoTo FIM
    End If
    
    Dim arrayItems2()
    With Planilha4
        ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, Sheets("Componentes").Range("A:A")), 1 To 8)
        Me.ListBox3.ColumnCount = 8
        Me.ListBox3.ColumnWidths = "40;130;350;80;90;50;90;50"
        i = 0
        For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
            Me.ListBox3.AddItem
            For coluna = 1 To 8
                arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
            Next coluna
            i = i + 1
        Next rngcell
        Me.ListBox3.List = arrayItems2()
    End With
FIM:
    Sheets("Componentes").Visible = xlSheetVeryHidden
End Sub
