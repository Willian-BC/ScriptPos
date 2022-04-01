Sub PECAS()
    Dim xPTable As PivotTable
    Dim xPFile As PivotField
    On Error Resume Next
    Application.ScreenUpdating = False
    Set xPTable = Worksheets("Dashboard").PivotTables("Tabela dinâmica1")
    Worksheets("Dashboard").xPTable.ManualUpdate = True
    
    xPTable.PivotFields("Demanda Kg").Orientation = xlHidden
    xPTable.PivotFields("Produzido (Kg)").Orientation = xlHidden
    xPTable.PivotFields("a Produzir (Kg)").Orientation = xlHidden
    xPTable.PivotFields("Expedido (Kg)").Orientation = xlHidden
    xPTable.PivotFields("a Expedir (Kg)").Orientation = xlHidden
    
    With xPTable.PivotFields("Demanda Pç")
        .Orientation = xlRowField
        .Position = 3
    End With
    With xPTable.PivotFields("Produzido (SKU)")
        .Orientation = xlRowField
        .Position = 4
    End With
    With xPTable.PivotFields("a Produzir (SKU)")
        .Orientation = xlRowField
        .Position = 6
    End With
    With xPTable.PivotFields("Expedido (SKU)")
        .Orientation = xlRowField
        .Position = 8
    End With
    With xPTable.PivotFields("a Expedir (SKU)")
        .Orientation = xlRowField
        .Position = 10
    End With
    
    Columns("O:V").HorizontalAlignment = xlCenter
    Columns("O:V").NumberFormat = "#,##0"
    Columns("P:P").NumberFormat = "#,##0.0"
    Columns("Q:Q").Style = "Percent"
    Columns("U:U").Style = "Percent"
    
    xPTable.ManualUpdate = False
    Application.ScreenUpdating = True
    
    ActiveSheet.Shapes("Retângulo: Cantos Arredondados 2").Fill.ForeColor.RGB = RGB(166, 166, 166)
    ActiveSheet.Shapes("Retângulo: Cantos Arredondados 3").Fill.ForeColor.RGB = RGB(166, 166, 166)
    ActiveSheet.Shapes("Retângulo: Cantos Arredondados 1").Fill.ForeColor.RGB = RGB(9, 71, 128)
End Sub
