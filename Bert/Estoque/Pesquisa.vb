Private Sub CommandButton1_Click()
'OK PESQUISA'
    Application.ScreenUpdating = False
    Sheets("Base de dados").Visible = xlSheetVisible
    Sheets("Base de dados").Activate
    Set wb = ActiveWorkbook
    
    If Not Sheets("Base de dados").AutoFilterMode Then
        Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
    End If
    On Error Resume Next
    Sheets("Base de dados").ShowAllData
    
    Set rngAF = Range("A1:A" & Cells(1, 1).End(xlDown).Row)
    
    If TextBox1 <> "" Then
        Sheets("Base de dados").Range("B:B").AutoFilter Field:=2, Criteria1:="=*" & TextBox1.Text & "*"
    End If
    If TextBox7 <> "" Then
        Sheets("Base de dados").Range("C:C").AutoFilter Field:=3, Criteria1:="=*" & TextBox7.Text & "*"
    End If
    If TextBox2 <> "" Then
        Sheets("Base de dados").Range("D:D").AutoFilter Field:=4, Criteria1:=TextBox2.Text
    End If
    If TextBox3 <> "" Then
        Sheets("Base de dados").Range("E:E").AutoFilter Field:=5, Criteria1:=TextBox3.Text
    End If
    If TextBox4 <> "" Then
        Sheets("Base de dados").Range("F:F").AutoFilter Field:=6, Criteria1:=TextBox4.Text
    End If
    If TextBox5 <> "" Then
        Sheets("Base de dados").Range("G:G").AutoFilter Field:=7, Criteria1:="=*" & TextBox5.Text & "*"
    End If
    If TextBox6 <> "" Then
        Sheets("Base de dados").Range("I:I").AutoFilter Field:=9, Criteria1:=TextBox6.Text
    End If
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = Sheets("Base de dados").Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        Sheets("Base de dados").ShowAllData
        GoTo FIM
    End If
    
'''''''''''''''''''''''CARREGAR LISTBOX'''''''''''''''''''''''''
    Dim arrayItems2()
    With Planilha5
        ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, Sheets("Base de dados").Range("A:A")), 1 To 11) '.UsedRange.Columns.Count
        Me.ListBox1.ColumnCount = 11 '.UsedRange.Columns.Count
        Me.ListBox1.ColumnWidths = "40;130;150;80;90;80;90;90;80;80;200"
        i = 0
        For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
            Me.ListBox1.AddItem
            For coluna = 1 To 11 '.UsedRange.Columns.Count
                arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
            Next coluna
            i = i + 1
        Next rngcell
        Me.ListBox1.List = arrayItems2()
    End With
    
    If CheckBox1 = True Or CheckBox2 = True Then

'''''''''''''''''''''''EXPORTAR DADOS'''''''''''''''''''''''''
        Set rngAJ = Range("A1:K" & lin_fim).SpecialCells(xlCellTypeVisible)
        rngAJ.Copy
        Workbooks.Add
        Range("A1").PasteSpecial Paste:=xlPasteValues
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:K" & Cells(1, 1).End(xlDown).Row), , xlYes).Name = "Tabela1"
        Columns("A:K").EntireColumn.AutoFit
'        Columns("K:K").NumberFormat = "dd/mm/yyyy"
        If CheckBox2 = True Then

'''''''''''''''''''''''IMPRIMIR DADOS'''''''''''''''''''''''''
            With Worksheets(1).PageSetup
            .Zoom = False
            .BlackAndWhite = True
            .FitToPagesTall = 1
            .FitToPagesWide = 1
            .CenterHorizontally = True
            .Orientation = xlLandscape
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.75)
            .BottomMargin = Application.InchesToPoints(0.75)
            .HeaderMargin = Application.InchesToPoints(0.3)
            .FooterMargin = Application.InchesToPoints(0.3)
            End With
            ActiveSheet.PrintOut
            MsgBox "IMPRESSÃO REALIZADA COM SUCESSO!", vbInformation
        End If
        If CheckBox1 = False Then
            ActiveWorkbook.Close SaveChanges:=False
        End If
    End If
    
FIM:
    wb.Sheets("Base de dados").Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
    
    If CheckBox1 = True Then
        result = MsgBox("DADOS EXPORTADOS COM SUCESSO !" & vbCrLf & "DESEJA FECHAR O FORMULÁRIO ?" & vbCrLf & vbCrLf & "É NECESSÁRIO FECHAR PARA EDITAR OS DADOS", vbYesNo + vbInformation)
        If result = vbYes Then
            Unload Me
        Else
'''''''''''''''''''''''MINIMIZAR PLANILHA CRIADA'''''''''''''''''''''''''
            ActiveWindow.WindowState = xlMinimized
            wb.Activate
            CheckBox1 = False
            CheckBox2 = False
        End If
    End If
    
End Sub
