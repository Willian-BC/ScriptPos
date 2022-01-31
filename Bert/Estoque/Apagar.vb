Private Sub CommandButton5_Click()
'APAGAR PESQUISA
    If ListBox1.ListCount = 0 Then
        MsgBox "FAVOR PREENCHER ALGUM CAMPO DE PESQUISA PARA CONTINUAR !", vbCritical
        GoTo FIM
    End If
    
    Application.ScreenUpdating = False
    Sheets("Base de dados").Visible = xlSheetVisible
    Sheets("Base de dados").Activate
    
    result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR UM REGISTRO?", vbYesNo + vbCritical)
    If result = vbYes Then
        ID = Application.InputBox("INFORME O ID")
        If ID = 0 Then
            GoTo FIM
        End If
        result = MsgBox("TEM CERTEZA QUE DESEJA EXCLUIR O ID " & ID & "?" & vbCrLf & vbCrLf & "ESSA AÇÃO NÃO É POSSÍVEL SER REVERTIDA", vbYesNo + vbCritical)
        If result = vbYes Then
            If Not Sheets("Base de dados").AutoFilterMode Then
                Sheets("Base de dados").Range("A1").End(xlToRight).AutoFilter
            End If
            On Error Resume Next
            Sheets("Base de dados").Range("A:A").AutoFilter Field:=1, Criteria1:=ID
            
            lin_inicio = Sheets("Base de dados").AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Row
            
            If Cells(lin_inicio, 1).Value = 0 Then
                MsgBox "VALOR INCORRETO !" & vbCrLf & vbCrLf & "FAVOR CONFIRMAR O NÚMERO DE ID", vbCritical
                Sheets("Base de dados").Range("A:A").AutoFilter Field:=1
                GoTo FIM
            End If
            
            Sheets("Base de dados").Rows(lin_inicio).Copy Sheets("Excluido").Rows(Sheets("Excluido").Cells(1, 1).End(xlDown).Row + 1)
            Sheets("Excluido").Cells(Sheets("Excluido").Cells(1, 1).End(xlDown).Row, 14) = Now
            Rows(lin_inicio).Delete
            
            Sheets("Base de dados").ShowAllData
            
            Sheets("Base de dados").Cells(2, 1) = 1
            Sheets("Base de dados").Cells(3, 1) = 2
            Sheets("Base de dados").Range("A2:A3").AutoFill Destination:=Range("A2:A" & Cells(1, 2).End(xlDown).Row)
            
            MsgBox "CADASTRO EXCLUIDO COM SUCESSO!", vbInformation
            ListBox1.Clear
        End If
    End If
FIM:
    Sheets("Base de dados").Visible = xlSheetVeryHidden
    Application.ScreenUpdating = True
End Sub
