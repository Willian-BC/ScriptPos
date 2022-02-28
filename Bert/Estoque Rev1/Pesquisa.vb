Public Sub CommandButton1_Click()
Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    Set ws = Sheets("Base de dados")
    
    If Not ws.AutoFilterMode Then ws.Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    ws.ShowAllData
    ws.Rows("2:" & ws.Cells(1, 1).End(xlDown).Row).ClearContents
    
    lsConectar
    Set lrs = New ADODB.Recordset
    lrs.Open " SELECT * FROM BD_dados ", gConexao, adOpenKeyset, adLockPessimistic
    ws.Cells(2, 1).CopyFromRecordset lrs
    ws.Columns("A:A").NumberFormat = "0"
    lrs.Close
    Set lrs = Nothing
    lsDesconectar
    
    Set rngAF = ws.Range("A1:A" & ws.Cells(1, 1).End(xlDown).Row)
    
    If TextBox1 <> "" Then ws.Range("B:B").AutoFilter Field:=2, Criteria1:="=*" & TextBox1.Text & "*"
    If TextBox2 <> "" Then ws.Range("C:C").AutoFilter Field:=3, Criteria1:="=*" & TextBox2.Text & "*"
    If TextBox3 <> "" Then ws.Range("D:D").AutoFilter Field:=4, Criteria1:=TextBox3.Text
    If TextBox4 <> "" Then ws.Range("F:F").AutoFilter Field:=6, Criteria1:=TextBox4.Text
    If TextBox5 <> "" Then ws.Range("G:G").AutoFilter Field:=7, Criteria1:="=*" & TextBox5.Text & "*"
    If TextBox6 <> "" Then ws.Range("E:E").AutoFilter Field:=5, Criteria1:=TextBox6.Text
    If TextBox7 <> "" Then ws.Range("I:I").AutoFilter Field:=9, Criteria1:=TextBox7.Text
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = ws.Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If ws.Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        ListBox1.Clear
        ws.ShowAllData
        GoTo FIM
    End If
    
    Dim arrayItems2()
    With Planilha5
        ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, ws.Range("A:A")), 1 To 11)
        Me.ListBox1.ColumnCount = 11
        Me.ListBox1.ColumnWidths = "40;120;300;70;70;70;200;70;70;70;200"
        i = 0
        For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
            Me.ListBox1.AddItem
            For coluna = 1 To 11
                arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
            Next coluna
            i = i + 1
        Next rngcell
        Me.ListBox1.List = arrayItems2()
    End With
    
    If CheckBox1 = True Or CheckBox2 = True Then
        Set rngAJ = ws.Range("A1:K" & lin_fim).SpecialCells(xlCellTypeVisible)
        rngAJ.Copy
        Workbooks.Add
        Range("A1").PasteSpecial Paste:=xlPasteValues
        ActiveSheet.ListObjects.Add(xlSrcRange, Range("A1:K" & Cells(1, 1).End(xlDown).Row), , xlYes).Name = "Tabela1"
        Columns("A:K").EntireColumn.AutoFit
        If CheckBox2 = True Then
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
        If CheckBox1 = False Then ActiveWorkbook.Close SaveChanges:=False
    End If
    If CheckBox1 = True Then
        result = MsgBox("DADOS EXPORTADOS COM SUCESSO !" & vbCrLf & "DESEJA FECHAR O FORMULÁRIO ?" & vbCrLf & vbCrLf & "É NECESSÁRIO FECHAR PARA EDITAR OS DADOS", vbYesNo + vbInformation)
        If result = vbYes Then
            Unload Me
        Else
            ActiveWindow.WindowState = xlMinimized
            wb.Activate
            CheckBox1 = False
            CheckBox2 = False
        End If
    End If
FIM:
    Application.ScreenUpdating = True
End Sub
