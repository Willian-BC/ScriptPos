Dim ws As Worksheet
Dim wb As Workbook

Private Sub MultiPage1_Change()

End Sub

Private Sub UserForm_Initialize()
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    Sheets("Baixa").Visible = xlSheetVeryHidden
    Sheets("Excluido").Visible = xlSheetVeryHidden
    Sheets("Descrição").Visible = xlSheetVeryHidden
    Sheets("Peso").Visible = xlSheetVeryHidden
End Sub

Private Sub CommandButton100_Click()
    Dim MyValue As Variant
    MyValue = InputBox("Digite a senha")
    If MyValue = "1010" Then
        Sheets("Base de dados").Visible = xlSheetVisible
        Sheets("Baixa").Visible = xlSheetVisible
        Sheets("Excluido").Visible = xlSheetVisible
        Sheets("Descrição").Visible = xlSheetVisible
        Sheets("Peso").Visible = xlSheetVisible
    Else
        MsgBox ("Senha Incorreta")
    End If
End Sub

Private Sub CommandButton200_Click()
    ThisWorkbook.Save
End Sub

Private Sub CommandButton300_Click()
    Dim OutlookApp As Object
    Dim OutlookMail As Object
    Dim ch1, ch2 As Chart
    
    Set OutlookApp = CreateObject("Outlook.Application")
    Set OutlookMail = OutlookApp.CreateItem(0)
    Set ch1 = ThisWorkbook.Worksheets("Descrição").ChartObjects("Gráfico 1").Chart
    Set ch2 = ThisWorkbook.Worksheets("Descrição").ChartObjects("Gráfico 2").Chart
    
    Const pic1 = "Chart1.png"
    Const pic2 = "Chart2.png"
    
    ch1.Export ThisWorkbook.Path & "\Chart1.png"
    ch2.Export ThisWorkbook.Path & "\Chart2.png"
    
    With OutlookMail
                .To = "renan.barros@bertolini.com.br;fellipe.novaes@bertolini.com.br;pcp.colatina@bertolini.com.br;margarete.comachio@bertolini.com.br;margarete.comachio@bertolini.com.br;danyele.rodrigues@bertolini.com.br;willian.martins@bertolini.com.br;italo.moschem@bertolini.com.br"
        .Subject = "Itens armazenados em estoque BMA"
        .Body = "Prezados," & vbNewLine
        .Attachments.Add ThisWorkbook.Path & "\" & pic1
        .Attachments.Add ThisWorkbook.Path & "\" & pic2
        .HTMLBody = "<html><p>Prezados,</p>" & _
                    "<p>Segue comparativo entre quantidade em estoque x quantidade máxima suportada:</p>" & _
                    "<img src=cid:" & Replace(pic1, " ", "%20") & " height=2*240 width=2*180>" & _
                    "<img src=cid:" & Replace(pic2, " ", "%20") & " height=2*240 width=2*180>" & _
                                                        "<p>Att," & _
                                                        "<p>" & Environ("USERNAME") & " - Expedição BMA" & "</p></html>"
'        .Display
        .Send
    End With
    
    Set OutlookMail = Nothing
    Set OutlookApp = Nothing
    
    Kill ThisWorkbook.Path & "\Chart1.png"
    Kill ThisWorkbook.Path & "\Chart2.png"
    
    MsgBox "EMAIL ENVIADO COM SUCESSO", vbOKOnly
    
End Sub
                        
Private Sub OptionButton3_Click()
    ListBox1.Clear
    TextBox6 = ""
    TextBox7 = ""
End Sub
Private Sub OptionButton4_Click()
    ListBox1.Clear
    TextBox6 = ""
    TextBox7 = ""
End Sub
Private Sub OptionButton5_Click()
    ListBox1.Clear
    TextBox6 = ""
    TextBox7 = ""
End Sub

Private Sub CommandButton1_Click()
'OK PESQUISA'
    
    If OptionButton3 = True Then
        Set ws = Sheets("Base de dados")
        a = "J:J"
        b = 10
    ElseIf OptionButton4 = True Then
        Set ws = Sheets("Baixa")
        a = "L:L"
        b = 12
    Else
        Set ws = Sheets("Excluido")
        a = "L:L"
        b = 12
    End If
    
    Application.ScreenUpdating = False
    Set wb = ActiveWorkbook
    
    If Not ws.AutoFilterMode Then
        ws.Range("A1").End(xlToRight).AutoFilter
    End If
    On Error Resume Next
    ws.ShowAllData
    ws.Columns("A:A,F:F").NumberFormat = "0"
    ws.Columns("G:G").NumberFormat = "#,##0.00"
    ws.Columns("J:L").NumberFormat = "dd/mm/yyyy"
    
    Set rngAF = ws.Range("A1:A" & ws.Cells(1, 1).End(xlDown).Row)
    
    If TextBox1 <> "" Then ws.Range("B:B").AutoFilter Field:=2, Criteria1:=TextBox1.Text
    If TextBox2 <> "" Then ws.Range("C:C").AutoFilter Field:=3, Criteria1:=TextBox2.Text
    If TextBox3 <> "" Then ws.Range("D:D").AutoFilter Field:=4, Criteria1:="=*" & TextBox3.Text & "*"
    If TextBox4 <> "" Then ws.Range("H:H").AutoFilter Field:=8, Criteria1:=TextBox4.Text
    If TextBox5 <> "" Then ws.Range(a).AutoFilter Field:=b, Criteria1:=Format(CDate(TextBox5), "dd/mm/yyyy")
    
    lin_inicio = ws.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    lin_fim = ws.Cells(1, 1).SpecialCells(xlCellTypeVisible).End(xlDown).Row
    
    If ws.Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALORES INFORMADOS NÃO EXISTEM !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR AS INFORMAÇÕES", vbCritical
        ws.ShowAllData
        GoTo FIM
    Else
    
    Dim arrayItems2()
        With ws
            ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, ws.Range("A:A")), 1 To 12)
            Me.ListBox1.ColumnCount = 12
            Me.ListBox1.ColumnWidths = "40;100;100;300;50;70;70;80;100;70;70;70;0"
            i = 0
            For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
                Me.ListBox1.AddItem
                For coluna = 1 To 12
                    arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
                Next coluna
                i = i + 1
            Next rngcell
            Me.ListBox1.List = arrayItems2()
        End With
    End If
    
    TextBox6 = Format(WorksheetFunction.Subtotal(109, ws.Range("F:F")), "0")
    TextBox7 = Format(WorksheetFunction.Subtotal(109, ws.Range("G:G")), "#,##0.00")
    
    If CheckBox1 = True Then
        Set rngAJ = ws.Range("B1:I" & lin_fim).SpecialCells(xlCellTypeVisible)
        rngAJ.Copy
        Workbooks.Add
        Range("A1").PasteSpecial Paste:=xlPasteValues
    End If
    
FIM:
    Application.ScreenUpdating = True
    
    If CheckBox1 = True Then
        result = MsgBox("DADOS EXPORTADOS COM SUCESSO !" & vbCrLf & "DESEJA FECHAR O FORMULÁRIO ?" & vbCrLf & vbCrLf & "É NECESSÁRIO FECHAR PARA EDITAR OS DADOS", vbYesNo + vbInformation)
        If result = vbYes Then
            Unload Me
        Else
            ActiveWindow.WindowState = xlMinimized
            wb.Activate
            CheckBox1 = False
        End If
    End If
    
End Sub

Private Sub CommandButton2_Click()
'LIMPAR PESQUISA
    TextBox1 = ""
    TextBox2 = ""
    TextBox3 = ""
    TextBox4 = ""
    TextBox5 = ""
    ListBox1.Clear
    CheckBox1 = False
    OptionButton3 = True
    OptionButton4 = False
    OptionButton5 = False
End Sub

Private Sub CommandButton3_Click()
'CANCELAR PESQUISA
    result = MsgBox("DESEJA SALVAR AS ALTERAÇÕES ?", vbYesNo + vbCritical)
    If result = vbYes Then ThisWorkbook.Save
    End
End Sub

Private Sub CommandButton4_Click()
'EDITAR PESQUISA
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    If OptionButton4 = True Then
        MsgBox "NÃO É POSSÍVEL EDITAR UM ITEM BAIXADO !" & vbCrLf & vbCrLf & "FAVOR ALTERAR O TIPO DE PESQUISA", vbCritical
        GoTo FIM
    ElseIf OptionButton5 = True Then
        MsgBox "NÃO É POSSÍVEL EDITAR UM ITEM EXCLUIDO !" & vbCrLf & vbCrLf & "FAVOR ALTERAR O TIPO DE PESQUISA", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
ERRO:
    ID = Application.InputBox("INFORME O ID")
    If ID = 0 Then GoTo FIM
    
    If Not Sheets("Base de dados").AutoFilterMode Then Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=ID
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    If Sheets("Base de dados").Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR O NÚMERO DE ID", vbCritical
        Sheets("Base de dados").Range("A:A").AutoFilter Field:=1
        GoTo ERRO
    End If
    
    TextBox101 = Sheets("Base de dados").Cells(lin_inicio, 1)
    TextBox102 = Sheets("Base de dados").Cells(lin_inicio, 2)
    TextBox103 = Sheets("Base de dados").Cells(lin_inicio, 3)
    TextBox104 = Sheets("Base de dados").Cells(lin_inicio, 4)
    larray = Split(Sheets("Base de dados").Cells(lin_inicio, 5), "/")
    TextBox1051 = larray(0)
    TextBox1052 = larray(1)
    TextBox106 = Sheets("Base de dados").Cells(lin_inicio, 6)
    TextBox107 = Sheets("Base de dados").Cells(lin_inicio, 7)
    TextBox108 = Sheets("Base de dados").Cells(lin_inicio, 8)
    TextBox109 = Sheets("Base de dados").Cells(lin_inicio, 9)
    
    Sheets("Base de dados").ShowAllData
    ListBox1.Clear
    UserForm1.MultiPage1.Value = 1
FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton5_Click()
'APAGAR PESQUISA
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    If OptionButton4 = True Then
        MsgBox "NÃO É POSSÍVEL DELETAR UM ITEM BAIXADO !" & vbCrLf & vbCrLf & "FAVOR ALTERAR O TIPO DE PESQUISA", vbCritical
        GoTo FIM
    ElseIf OptionButton5 = True Then
        MsgBox "NÃO É POSSÍVEL DELETAR UM ITEM EXCLUIDO !" & vbCrLf & vbCrLf & "FAVOR ALTERAR O TIPO DE PESQUISA", vbCritical
        GoTo FIM
    End If
    
ERRO:
    ID = Application.InputBox("INFORME O ID")
    If ID = 0 Then GoTo FIM
    
    result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR O ID " & ID & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
    If result = vbYes Then
        
        Application.ScreenUpdating = False
        If Not Sheets("Base de dados").AutoFilterMode Then Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
        On Error Resume Next
        Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=ID
        
        lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
        
        If Sheets("Base de dados").Cells(lin_inicio, 1).Value = 0 Then
            MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR O NÚMERO DE ID", vbCritical
            Sheets("Base de dados").Range("A:A").AutoFilter Field:=1
            GoTo ERRO
        End If
        
        Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Excluido").Rows(Sheets("Excluido").Cells(1, 1).End(xlDown).Row + 1)
        Sheets("Excluido").Cells(Sheets("Excluido").Cells(1, 1).End(xlDown).Row, 12) = Now
        Sheets("Base de dados").Rows(lin_inicio).Delete
        
        Sheets("Base de dados").ShowAllData
        
        Sheets("Base de dados").Cells(2, 1) = 1
        Sheets("Base de dados").Cells(3, 1) = 2
        Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Sheets("Base de dados").Range("A2:A" & Sheets("Base de dados").Cells(1, 1).End(xlDown).Row)
        
        MsgBox "CADASTRO EXCLUIDO COM SUCESSO!", vbInformation
        ListBox1.Clear
    End If
FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton101_Click()
'OK CADASTRO'
    If TextBox101 <> "" Then
        MsgBox "CADASTRO JÁ EXISTE!" & vbCrLf & vbCrLf & "CLIQUE EM ATUALIZAR REGISTRO OU LIMPAR", vbCritical
        GoTo FIM
    ElseIf TextBox102 = "" Or TextBox106 = "" Or TextBox108 = "" Then
        MsgBox "FAVOR PREENCHER TODOS OS CAMPOS PARA CADASTRO!", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    
    On Error Resume Next
    Sheets("Base de dados").ShowAllData
    lin = Sheets("Base de dados").Cells(1, 1).End(xlDown).Row + 1
    Sheets("Base de dados").Cells(lin, 1) = lin - 1
    Sheets("Base de dados").Cells(lin, 2) = TextBox102.Text
    Sheets("Base de dados").Cells(lin, 3) = TextBox103.Text
    Sheets("Base de dados").Cells(lin, 4) = TextBox104.Text
    Sheets("Base de dados").Cells(lin, 5) = TextBox1051.Text & "/" & TextBox1052.Text
    Sheets("Base de dados").Cells(lin, 6) = TextBox106.Text
    Sheets("Base de dados").Cells(lin, 7) = TextBox107.Text
    Sheets("Base de dados").Cells(lin, 8) = TextBox108.Text
    Sheets("Base de dados").Cells(lin, 9) = TextBox109.Text
    Sheets("Base de dados").Cells(lin, 10) = Now
    
    result = MsgBox("CADASTRO REALIZADO COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA LIMPAR OS DADOS?", vbYesNo + vbInformation)
    If result = vbYes Then
        CommandButton102_Click
    End If
    
    TextBox102.SetFocus
    
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub CommandButton102_Click()
'LIMPAR CADASTRO
    TextBox101 = ""
    TextBox102 = ""
    TextBox103 = ""
    TextBox104 = ""
    TextBox1051 = ""
    TextBox1052 = ""
    TextBox106 = ""
    TextBox107 = ""
    TextBox108 = ""
    TextBox109 = ""
End Sub

Private Sub CommandButton103_Click()
'CANCELAR CADASTRO
    result = MsgBox("DESEJA SALVAR AS ALTERAÇÕES ?", vbYesNo + vbCritical)
    If result = vbYes Then ThisWorkbook.Save
    End
End Sub

Private Sub CommandButton104_Click()
'ATUALIZAR CADASTRO
    If TextBox101 = "" Then
        MsgBox "CADASTRO NÃO EXISTE!" & vbCrLf & vbCrLf & "CLIQUE EM CADASTRAR", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    
    If Not Sheets("Base de dados").AutoFilterMode Then Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    Sheets("Base de dados").ShowAllData
    Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=TextBox101.Text
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    Sheets("Base de dados").Cells(lin_inicio, 2) = TextBox102.Text
    Sheets("Base de dados").Cells(lin_inicio, 3) = TextBox103.Text
    Sheets("Base de dados").Cells(lin_inicio, 4) = TextBox104.Text
    Sheets("Base de dados").Cells(lin_inicio, 5) = TextBox1051.Text & "/" & TextBox1052.Text
    Sheets("Base de dados").Cells(lin_inicio, 6) = TextBox106.Text
    Sheets("Base de dados").Cells(lin_inicio, 7) = TextBox107.Text
    Sheets("Base de dados").Cells(lin_inicio, 8) = TextBox108.Text
    Sheets("Base de dados").Cells(lin_inicio, 9) = TextBox109.Text
    Sheets("Base de dados").Cells(lin_inicio, 11) = Now
    
    Sheets("Base de dados").ShowAllData
    
    result = MsgBox("CADASTRO ATUALIZADO COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA LIMPAR OS DADOS?", vbYesNo + vbInformation)
    If result = vbYes Then
        CommandButton102_Click
        UserForm1.MultiPage1.Value = 0
    End If
    Application.ScreenUpdating = True
FIM:
End Sub

Private Sub CommandButton201_Click()
'OK BAIXA

    If TextBox201 = "" And TextBox202 = "" And TextBox204 = "" Then
        MsgBox "FAVOR PREENCHER AS INFORMAÇÕES DE BAIXA!", vbCritical
        GoTo FIM
    End If
    If TextBox204 <> "" And OptionButton1.Value = False And OptionButton2.Value = False Then
        MsgBox "FAVOR PREENCHER O TIPO DE BAIXA!", vbCritical
        GoTo FIM
    End If
    Application.ScreenUpdating = False
    If Not Sheets("Base de dados").AutoFilterMode Then Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
NOVO:
    Sheets("Base de dados").ShowAllData
    Set rngAF = Sheets("Base de dados").Range("A1:A" & Sheets("Base de dados").Cells(1, 1).End(xlDown).Row)
    
    If TextBox201 <> "" Then Sheets("Base de dados").Range("B:B").AutoFilter Field:=2, Criteria1:=TextBox201.Text
    If TextBox202 <> "" Then Sheets("Base de dados").Range("C:C").AutoFilter Field:=3, Criteria1:=TextBox202.Text
    If TextBox204 <> "" Then Sheets("Base de dados").Range("H:H").AutoFilter Field:=8, Criteria1:=TextBox204.Text
    
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    If Sheets("Base de dados").Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALOR INFORMADO NÃO EXISTE !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR INFORMAÇÕES", vbCritical
        GoTo FIM
    Else
        Dim arrayItems2()
        With Planilha5
            ReDim arrayItems2(0 To WorksheetFunction.Subtotal(102, Sheets("Base de dados").Range("A:A")), 1 To 9)
            Me.ListBox2.ColumnCount = 9
            Me.ListBox2.ColumnWidths = "40;100;100;300;50;70;70;80;100"
            i = 0
            For Each rngcell In rngAF.SpecialCells(xlCellTypeVisible)
                Me.ListBox2.AddItem
                For coluna = 1 To 9
                    arrayItems2(i, coluna) = .Cells(rngcell.Row, coluna).Value
                Next coluna
                i = i + 1
            Next rngcell
            Me.ListBox2.List = arrayItems2()
        End With
    End If
    
    If OptionButton1.Value = True Then
        result = MsgBox("TEM CERTEZA QUE DESEJA DAR BAIXA TODOS OS ITENS DA POSIÇÃO " & TextBox204 & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
        If result = vbYes Then
            Do While Sheets("Base de dados").Cells(lin_inicio, 1).Value <> 0
                Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Baixa").Rows(Sheets("Baixa").Cells(1, 1).End(xlDown).Row + 1)
                Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 12) = Now
                Sheets("Base de dados").Rows(lin_inicio).Delete
                lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            Loop
            Sheets("Base de dados").ShowAllData
        
            Sheets("Base de dados").Cells(2, 1) = 1
            Sheets("Base de dados").Cells(3, 1) = 2
            Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Sheets("Base de dados").Range("A2:A" & Sheets("Base de dados").Cells(1, 1).End(xlDown).Row)
        Else
            Sheets("Base de dados").ShowAllData
            GoTo FIM
        End If
        MsgBox "BAIXA REALIZADA COM SUCESSO !", vbInformation
        CommandButton202_Click
        GoTo FIM
    End If
ERRO2:
    ID = Application.InputBox("INFORME O ID")
    If ID = 0 Then
        Sheets("Base de dados").ShowAllData
        Sheets("Base de dados").Cells(2, 1) = 1
        Sheets("Base de dados").Cells(3, 1) = 2
        Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Sheets("Base de dados").Range("A2:A" & Sheets("Base de dados").Cells(1, 1).End(xlDown).Row)
        GoTo FIM
    End If
    
    If Not Sheets("Base de dados").AutoFilterMode Then Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
    On Error Resume Next
    Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=ID
    lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
    
    If Sheets("Base de dados").Cells(lin_inicio, 1).Value = 0 Then
        MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR O NÚMERO DE ID", vbCritical
        Sheets("Base de dados").Range("A:A").AutoFilter Field:=1
        GoTo ERRO2
    End If
    
ERRO:
    result = Application.InputBox("INFORME A QUANTIDADE QUE DESEJA DAR BAIXA NO ID " & ID & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA")
    If result > 0 Then
        If Sheets("Base de dados").Cells(lin_inicio, 6).Value - result > 0 Then
            Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Baixa").Rows(Sheets("Baixa").Cells(1, 1).End(xlDown).Row + 1)
            Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 6) = result
            Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 12) = Now
            Sheets("Base de dados").Cells(lin_inicio, 6).Value = Sheets("Base de dados").Cells(lin_inicio, 6).Value - result
        ElseIf Sheets("Base de dados").Cells(lin_inicio, 6).Value - result = 0 Then
            Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Baixa").Rows(Sheets("Baixa").Cells(1, 1).End(xlDown).Row + 1)
            Sheets("Baixa").Cells(Sheets("Baixa").Cells(1, 1).End(xlDown).Row, 12) = Now
            Sheets("Base de dados").Rows(lin_inicio).Delete
        Else
            MsgBox "VALOR INFORMADO MAIOR QUE ESTOQUE" & vbCrLf & vbCrLf & "VERIFIQUE A QUANTIDADE INFORMADA"
            GoTo ERRO
        End If
        
        Sheets("Base de dados").ShowAllData
        result = MsgBox("BAIXA REALIZADA COM SUCESSO!" & vbCrLf & vbCrLf & "DESEJA BAIXAR OUTRO ITEM DESSE PEDIDO ?", vbYesNo + vbInformation)
        If result = vbYes Then
            GoTo NOVO
        Else
            Sheets("Base de dados").Cells(2, 1) = 1
            Sheets("Base de dados").Cells(3, 1) = 2
            Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Sheets("Base de dados").Range("A2:A" & Sheets("Base de dados").Cells(1, 1).End(xlDown).Row)
        End If
    End If

FIM:
    Application.ScreenUpdating = True
End Sub

Private Sub CommandButton202_Click()
'LIMPAR BAIXA
    TextBox201 = ""
    TextBox202 = ""
    TextBox204 = ""
    ListBox2.Clear
    OptionButton1.Value = False
    OptionButton2.Value = False
    OptionButton1.Enabled = False
    OptionButton2.Enabled = False
End Sub

Private Sub CommandButton203_Click()
'CANCELAR BAIXA
    result = MsgBox("DESEJA SALVAR AS ALTERAÇÕES ?", vbYesNo + vbCritical)
    If result = vbYes Then ThisWorkbook.Save
    End
End Sub

Private Sub TextBox1_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox2_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox3_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox4_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If IsNumeric(Left(TextBox4, 1)) Then
        TextBox4.MaxLength = 8
        Select Case KeyAscii
            Case 8 'Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}" 'Emula o TAB
            Case 48 To 57
            If TextBox4.SelStart = 2 Then
                TextBox4.SelText = "-"
            End If
            If TextBox4.SelStart = 5 Then
                TextBox4.SelText = "-"
            End If
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
        End Select
    Else
        TextBox4.MaxLength = 7
        Select Case KeyAscii
            Case 8 'Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}" 'Emula o TAB
            Case 48 To 90
            If TextBox4.SelStart = 1 Then
                TextBox4.SelText = "-"
            End If
            If TextBox4.SelStart = 4 Then
                TextBox4.SelText = "-"
            End If
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
        End Select
    End If
End Sub

Private Sub TextBox5_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    TextBox5.MaxLength = 10
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        If TextBox5.SelStart = 2 Then
            TextBox5.SelText = "/"
        End If
        If TextBox5.SelStart = 5 Then
            TextBox5.SelText = "/"
        End If
    Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox102_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox102_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If TextBox101 = "" Then
        If TextBox102 <> "" Then
            Sheets("Descrição").Cells(1, 4) = TextBox102.Text
            TextBox104 = UCase(Worksheets("Descrição").Cells(1, 5).Value)
        Else
            TextBox104 = ""
        End If
        TextBox107 = ""
    End If
End Sub

Private Sub TextBox103_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox104_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox1051_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox1052_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox106_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox106_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    If TextBox101 = "" Then
        If TextBox102 <> "" And TextBox106 <> "" Then
            Sheets("Peso").Cells(2, 9) = TextBox102.Text
            TextBox107 = Format(TextBox106 * Worksheets("Peso").Cells(3, 9).Value, "#,##0.00")
        Else
            TextBox107 = ""
        End If
    End If
End Sub

Private Sub TextBox107_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case 8 'Aceita o BACK SPACE
        Case 13: SendKeys "{TAB}" 'Emula o TAB
        Case 48 To 57
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
    End Select
End Sub

Private Sub TextBox108_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If IsNumeric(Left(TextBox108, 1)) Then
        TextBox108.MaxLength = 8
        Select Case KeyAscii
            Case 8 'Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}" 'Emula o TAB
            Case 48 To 57
            If TextBox108.SelStart = 2 Then
                TextBox108.SelText = "-"
            End If
            If TextBox108.SelStart = 5 Then
                TextBox108.SelText = "-"
            End If
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
        End Select
    Else
        TextBox108.MaxLength = 7
        Select Case KeyAscii
            Case 8 'Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}" 'Emula o TAB
            Case 48 To 90
            If TextBox108.SelStart = 1 Then
                TextBox108.SelText = "-"
            End If
            If TextBox108.SelStart = 4 Then
                TextBox108.SelText = "-"
            End If
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
        End Select
    End If
End Sub

Private Sub TextBox109_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub TextBox204_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If IsNumeric(Left(TextBox204, 1)) Then
        TextBox204.MaxLength = 8
        Select Case KeyAscii
            Case 8 'Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}" 'Emula o TAB
            Case 48 To 57
            If TextBox204.SelStart = 2 Then
                TextBox204.SelText = "-"
            End If
            If TextBox204.SelStart = 5 Then
                TextBox204.SelText = "-"
            End If
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
        End Select
    Else
        TextBox204.MaxLength = 7
        Select Case KeyAscii
            Case 8 'Aceita o BACK SPACE
            Case 13: SendKeys "{TAB}" 'Emula o TAB
            Case 48 To 90
            If TextBox204.SelStart = 1 Then
                TextBox204.SelText = "-"
            End If
            If TextBox204.SelStart = 4 Then
                TextBox204.SelText = "-"
            End If
        Case Else: KeyAscii = 0 'Ignora os outros caracteres
        End Select
    End If
End Sub

Private Sub TextBox204_Change()
    If IsNumeric(Left(TextBox204, 1)) And TextBox204 <> "" And Len(TextBox204) = 8 Then
        OptionButton1.Enabled = True
        OptionButton2.Enabled = True
    ElseIf Not (IsNumeric(Left(TextBox204, 1))) And TextBox204 <> "" And Len(TextBox204) = 7 Then
        OptionButton1.Enabled = True
        OptionButton2.Enabled = True
    Else
        OptionButton1.Enabled = False
        OptionButton2.Enabled = False
        OptionButton1.Value = False
        OptionButton2.Value = False
    End If
End Sub
