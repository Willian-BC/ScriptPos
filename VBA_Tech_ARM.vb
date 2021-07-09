Sub Macro1()
    ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 1")).Select
    Selection.ShapeRange.IncrementLeft -0.75
    Selection.ShapeRange.IncrementTop 9
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset10
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset11
    Selection.ShapeRange.ShapeStyle = msoShapeStylePreset10
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "OK"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 2). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignLeft
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 2).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Name = "+mn-lt"
    End With
    Selection.ShapeRange.IncrementLeft 6.75
    Selection.ShapeRange.IncrementTop -3.75
    Selection.ShapeRange(1).TextFrame2.TextRange.Characters.Text = "NÃO OK"
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6). _
        ParagraphFormat
        .FirstLineIndent = 0
        .Alignment = msoAlignLeft
    End With
    With Selection.ShapeRange(1).TextFrame2.TextRange.Characters(1, 6).Font
        .NameComplexScript = "+mn-cs"
        .NameFarEast = "+mn-ea"
        .Fill.Visible = msoTrue
        .Fill.ForeColor.ObjectThemeColor = msoThemeColorLight1
        .Fill.ForeColor.TintAndShade = 0
        .Fill.ForeColor.Brightness = 0
        .Fill.Transparency = 0
        .Fill.Solid
        .Size = 11
        .Name = "+mn-lt"
    End With
    Range("A1:J1").Select
End Sub

Sub BOTAO_OK()
   'G1
    If WorksheetFunction.CountA(Range("C8:D19")) = 24 Then
        For i = 8 To 19
            If (Cells(i, 3).Value <= Cells(5, 6).Value And Cells(i, 3).Value >= Cells(4, 6).Value _
            And Cells(i, 4).Value >= Cells(4, 7).Value And Cells(i, 4).Value <= Cells(6, 7).Value) Or _
            (Cells(i, 3).Value <= Cells(6, 6).Value And Cells(i, 3).Value >= Cells(5, 6).Value _
            And Cells(i, 4).Value >= Cells(4, 7).Value And Cells(i, 4).Value <= Cells(7, 7).Value) Or _
            (Cells(i, 3).Value <= Cells(7, 6).Value And Cells(i, 3).Value >= Cells(6, 6).Value _
            And Cells(i, 4).Value >= Cells(5, 7).Value And Cells(i, 4).Value <= Cells(7, 7).Value) Then
                green1 = green1 + 1
                GoTo PROX1
            ElseIf Cells(i, 3).Value >= WorksheetFunction.Average(Range("C8:C19")) * 1.2 Or Cells(i, 3).Value <= WorksheetFunction.Average(Range("C8:C19")) * 0.8 _
            Or Cells(i, 4).Value >= WorksheetFunction.Average(Range("D8:D19")) * 1.2 Or Cells(i, 4).Value <= WorksheetFunction.Average(Range("D8:D19")) * 0.8 Then
                red1 = red1 + 1
                GoTo PROX1
            ElseIf (WorksheetFunction.Average(Range("C8:C19")) <= Cells(5, 6).Value And WorksheetFunction.Average(Range("C8:C19")) >= Cells(4, 6).Value _
            And WorksheetFunction.Average(Range("D8:D19")) >= Cells(4, 7).Value And WorksheetFunction.Average(Range("D8:D19")) <= Cells(6, 7).Value _
            And Cells(i, 3).Value <= WorksheetFunction.Average(Range("C8:C19")) * 1.2 And Cells(i, 3).Value >= WorksheetFunction.Average(Range("C8:C19")) * 0.8 _
            And Cells(i, 4).Value <= WorksheetFunction.Average(Range("D8:D19")) * 1.2 And Cells(i, 4).Value >= WorksheetFunction.Average(Range("D8:D19")) * 0.8) Or _
            (WorksheetFunction.Average(Range("C8:C19")) <= Cells(6, 6).Value And WorksheetFunction.Average(Range("C8:C19")) >= Cells(5, 6).Value _
            And WorksheetFunction.Average(Range("D8:D19")) >= Cells(4, 7).Value And WorksheetFunction.Average(Range("D8:D19")) <= Cells(7, 7).Value _
            And Cells(i, 3).Value <= WorksheetFunction.Average(Range("C8:C19")) * 1.2 And Cells(i, 3).Value >= WorksheetFunction.Average(Range("C8:C19")) * 0.8 _
            And Cells(i, 4).Value <= WorksheetFunction.Average(Range("D8:D19")) * 1.2 And Cells(i, 4).Value >= WorksheetFunction.Average(Range("D8:D19")) * 0.8) Or _
            (WorksheetFunction.Average(Range("C8:C19")) <= Cells(7, 6).Value And WorksheetFunction.Average(Range("C8:C19")) >= Cells(6, 6).Value _
            And WorksheetFunction.Average(Range("D8:D19")) >= Cells(5, 7).Value And WorksheetFunction.Average(Range("D8:D19")) <= Cells(7, 7).Value _
            And Cells(i, 3).Value <= WorksheetFunction.Average(Range("C8:C19")) * 1.2 And Cells(i, 3).Value >= WorksheetFunction.Average(Range("C8:C19")) * 0.8 _
            And Cells(i, 4).Value <= WorksheetFunction.Average(Range("D8:D19")) * 1.2 And Cells(i, 4).Value >= WorksheetFunction.Average(Range("D8:D19")) * 0.8) Then
                yellow1 = 1
                GoTo PROX1
'            ElseIf (Cells(i, 3).Value <= WorksheetFunction.Average(Range("C8:C19")) * 1.2 _
'            And Cells(i, 3).Value >= WorksheetFunction.Average(Range("C8:C19")) * 0.8) Or _
'            (Cells(i, 4).Value <= WorksheetFunction.Average(Range("D8:D19")) * 1.2 _
'            And Cells(i, 4).Value >= WorksheetFunction.Average(Range("D8:D19")) * 0.8) Then
            End If
PROX1:
        Next
        If green1 = 12 Then
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 1")).Visible = True
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 1")).ShapeStyle = msoShapeStylePreset11
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 1")).TextFrame2.TextRange.Characters.Text = "SEGUIR PRODUÇÃO"
        ElseIf red1 > 0 Then
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 1")).Visible = True
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 1")).ShapeStyle = msoShapeStylePreset10
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 1")).TextFrame2.TextRange.Characters.Text = "CORRIGIR DADOS"
            MsgBox ("EXISTE(M) " & red1 & " CABEÇA(S) FORA DO PADRÃO NA G1")
            End
        ElseIf yellow1 = 1 Then
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 1")).Visible = True
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 1")).ShapeStyle = msoShapeStylePreset14
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 1")).TextFrame2.TextRange.Characters.Text = "AVALIAR DADOS"
        End If
    Else
        ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 1")).Visible = False
    End If
    'G2
    If WorksheetFunction.CountA(Range("C22:D33")) = 24 Then
        For i = 22 To 33
            If (Cells(i, 3).Value <= Cells(5, 9).Value And Cells(i, 3).Value >= Cells(4, 9).Value _
            And Cells(i, 4).Value >= Cells(4, 10).Value And Cells(i, 4).Value <= Cells(6, 10).Value) Or _
            (Cells(i, 3).Value <= Cells(6, 9).Value And Cells(i, 3).Value >= Cells(5, 9).Value _
            And Cells(i, 4).Value >= Cells(4, 10).Value And Cells(i, 4).Value <= Cells(7, 10).Value) Or _
            (Cells(i, 3).Value <= Cells(7, 9).Value And Cells(i, 3).Value >= Cells(6, 9).Value _
            And Cells(i, 4).Value >= Cells(5, 10).Value And Cells(i, 4).Value <= Cells(7, 10).Value) Then
                green2 = green2 + 1
                GoTo PROX2
            ElseIf Cells(i, 3).Value >= WorksheetFunction.Average(Range("C22:C33")) * 1.2 Or Cells(i, 3).Value <= WorksheetFunction.Average(Range("C22:C33")) * 0.8 _
            Or Cells(i, 4).Value >= WorksheetFunction.Average(Range("D22:D33")) * 1.2 Or Cells(i, 4).Value <= WorksheetFunction.Average(Range("D22:D33")) * 0.8 Then
                red2 = red2 + 1
                GoTo PROX2
            ElseIf (WorksheetFunction.Average(Range("C22:C33")) <= Cells(5, 9).Value And WorksheetFunction.Average(Range("C22:C33")) >= Cells(4, 9).Value _
            And WorksheetFunction.Average(Range("D22:D33")) >= Cells(4, 10).Value And WorksheetFunction.Average(Range("D22:D33")) <= Cells(6, 10).Value _
            And Cells(i, 3).Value <= WorksheetFunction.Average(Range("C22:C33")) * 1.2 And Cells(i, 3).Value >= WorksheetFunction.Average(Range("C22:C33")) * 0.8 _
            And Cells(i, 4).Value <= WorksheetFunction.Average(Range("D22:D33")) * 1.2 And Cells(i, 4).Value >= WorksheetFunction.Average(Range("D22:D33")) * 0.8) Or _
            (WorksheetFunction.Average(Range("C22:C33")) <= Cells(6, 9).Value And WorksheetFunction.Average(Range("C22:C33")) >= Cells(5, 9).Value _
            And WorksheetFunction.Average(Range("D22:D33")) >= Cells(4, 10).Value And WorksheetFunction.Average(Range("D22:D33")) <= Cells(7, 10).Value _
            And Cells(i, 3).Value <= WorksheetFunction.Average(Range("C22:C33")) * 1.2 And Cells(i, 3).Value >= WorksheetFunction.Average(Range("C22:C33")) * 0.8 _
            And Cells(i, 4).Value <= WorksheetFunction.Average(Range("D22:D33")) * 1.2 And Cells(i, 4).Value >= WorksheetFunction.Average(Range("D22:D33")) * 0.8) Or _
            (WorksheetFunction.Average(Range("C22:C33")) <= Cells(7, 9).Value And WorksheetFunction.Average(Range("C22:C33")) >= Cells(6, 9).Value _
            And WorksheetFunction.Average(Range("D22:D33")) >= Cells(5, 10).Value And WorksheetFunction.Average(Range("D22:D33")) <= Cells(7, 10).Value _
            And Cells(i, 3).Value <= WorksheetFunction.Average(Range("C22:C33")) * 1.2 And Cells(i, 3).Value >= WorksheetFunction.Average(Range("C22:C33")) * 0.8 _
            And Cells(i, 4).Value <= WorksheetFunction.Average(Range("D22:D33")) * 1.2 And Cells(i, 4).Value >= WorksheetFunction.Average(Range("D22:D33")) * 0.8) Then
                yellow2 = 1
                GoTo PROX2
            End If
PROX2:
        Next
        If green2 = 12 Then
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 2")).Visible = True
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 2")).ShapeStyle = msoShapeStylePreset11
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 2")).TextFrame2.TextRange.Characters.Text = "SEGUIR PRODUÇÃO"
        ElseIf red2 > 0 Then
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 2")).Visible = True
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 2")).ShapeStyle = msoShapeStylePreset10
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 2")).TextFrame2.TextRange.Characters.Text = "CORRIGIR DADOS"
            MsgBox ("EXISTE(M) " & red2 & " CABEÇA(S) FORA DO PADRÃO NA G2")
            End
        ElseIf yellow2 = 1 Then
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 2")).Visible = True
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 2")).ShapeStyle = msoShapeStylePreset14
            ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 2")).TextFrame2.TextRange.Characters.Text = "AVALIAR DADOS"
        End If
    Else
        ActiveSheet.Shapes.Range(Array("Snip Single Corner Rectangle 2")).Visible = False
    End If
End Sub

Sub salva()
    If Range("E2").Value <> "" Then
        savename = Range("E2").Value & " " & Format(Now, "dd-mm-yy hh-mm-ss")
        If Sheets(1).CheckBox5 = True Then
            Sheets(1).Unprotect Password:="pmb3210"
            Cells(1, 7).Value = Date
            Sheets(1).Protect Password:="pmb3210", DrawingObjects:=True, Contents:=True, Scenarios:=True
            ActiveWorkbook.SaveCopyAs "M:\VIX_QHSE\VIX_QUALIDADE\P_PUBLIC\P25_ FICHA DE AVALIACAO DA REGULAGEM DA ARMAGEM\" _
            & savename & "_A1" & ".xlsm"
            ActiveWorkbook.SaveCopyAs "M:\VIX_PROCESSOS\PUBLIC\3 - TENSAO E DEFORMACAO NOS FIOS\08 - Backup dos formulários de puxada de fio\" _
            & savename & "_A1" & ".xlsm"
            MsgBox ("DOCUMENTO SALVO COM SUCESSO!!"), vbInformation
        ElseIf Sheets(1).CheckBox6 = True Then
            Sheets(1).Unprotect Password:="pmb3210"
            Cells(1, 7).Value = Date
            Sheets(1).Protect Password:="pmb3210", DrawingObjects:=True, Contents:=True, Scenarios:=True
            ActiveWorkbook.SaveCopyAs "M:\VIX_QHSE\VIX_QUALIDADE\P_PUBLIC\P25_ FICHA DE AVALIACAO DA REGULAGEM DA ARMAGEM\" _
            & savename & "_A2" & ".xlsm"
            ActiveWorkbook.SaveCopyAs "M:\VIX_PROCESSOS\PUBLIC\3 - TENSAO E DEFORMACAO NOS FIOS\08 - Backup dos formulários de puxada de fio\" _
            & savename & "_A2" & ".xlsm"
            MsgBox ("DOCUMENTO SALVO COM SUCESSO!!"), vbInformation
        Else
            MsgBox ("PRENCHER CAMPO DE FASE"), vbExclamation
        End If
    Else
        MsgBox ("PRENCHER CAMPO LANÇAMENTO!!"), vbExclamation
    End If
End Sub

Sub limpar()

    CLEAR.Show
    
End Sub

' Userform1 Limpeza dos dados

Private Sub CommandButton1_Click()
    Sheet1.Range("B8:I19").Value = ""
    Sheet1.Range("B22:I33").Value = ""
    Sheet1.Range("B2:B5").Value = ""
    Sheet1.Range("D3:D6").Value = ""
    Sheet1.Range("E2").Value = ""
    Sheet1.Range("G2").Value = ""
    Sheet1.Range("I1:J2").Value = ""
    Sheet1.Range("D20").Value = ""
    Sheet1.Range("C35").Value = ""
    Sheet1.Range("F35").Value = ""
    Sheet1.Range("H35").Value = ""
    Sheet1.CheckBox1.Value = False
    Sheet1.CheckBox2.Value = False
    Sheet1.CheckBox3.Value = False
    Sheet1.CheckBox4.Value = False
    Sheet1.CheckBox5.Value = False
    Sheet1.CheckBox6.Value = False
    Sheet1.CheckBox7.Value = False
    Sheet1.CheckBox8.Value = False
    End
End Sub

Private Sub CommandButton2_Click()
    End
End Sub

Private Sub OptionButton1_Click()
    CommandButton1.Enabled = True
End Sub
