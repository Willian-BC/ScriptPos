'DASHBOARD
Private Sub Worksheet_Change(ByVal Target As Range)
'Update by Extendoffice 20180702
    Application.Calculation = xlAutomatic
    Dim xPTable As PivotTable
    Dim xPFile As PivotField
    Dim xStr As String
    On Error Resume Next
    If Intersect(Target, Range("E2")) Is Nothing Then Exit Sub
    Application.ScreenUpdating = False
    Set xPTable = Worksheets("Dashboard").PivotTables("Tabela dinâmica1")
    Set xPFile = xPTable.PivotFields("Pedido")
    xStr = Target.Text
    xPFile.ClearAllFilters
    xPFile.CurrentPage = xStr
    Application.ScreenUpdating = True
End Sub
      
'PESQUISA
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim pvt As PivotTable
    Dim pvtField As PivotField
    On Error Resume Next
    'If ActiveSheet.Cells(2, 4) = "" Then Exit Sub
    Application.ScreenUpdating = False
    Application.Calculation = xlAutomatic
    Set pvt = ActiveSheet.PivotTables("Tabela dinâmica2")
    Set pvtField = pvt.PivotFields("Fantasia")
    pvtField.ClearAllFilters
    pvtField.PivotFilters.Add xlCaptionEquals, Value1:="*" & Cells(2, 4).Value & "*"
    Application.ScreenUpdating = True
End Sub
