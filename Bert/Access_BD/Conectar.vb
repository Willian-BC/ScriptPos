Option Explicit
Dim gConexao As New ADODB.Connection
Dim lrs As New ADODB.recordset
Dim BD_command As New ADODB.Command
Dim strConexao, tabela, sql, nID As String
Dim ws As Worksheet

Private Sub lsConectar()
    Set gConexao = New ADODB.Connection
    
    strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=H:\Grupos\Database_EXP.accdb;Persist Security Info=False"
    gConexao.Open strConexao
    
    If gConexao.State = adStateOpen Then
        MsgBox " Conexão ativa "
    ElseIf gConexao.State = adStateClosed Then
        MsgBox " Conexão falhou tente novamente", vbCritical
    End If
    
End Sub
