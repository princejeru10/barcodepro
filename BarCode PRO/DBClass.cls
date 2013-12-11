VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim cmd As New ADODB.Command
Dim connString As String
Option Explicit

Public Sub dbOpen()
    'conn.ConnectionString = "Provider=SQLNCLI10;Server=ACCTG-CHRIS\SQLEXPRESS;DataTypeCompatibility=80;Database=BCBeta;User Id=carlo_;Password=/stats;"
    conn.ConnectionString = "Driver={MySQL ODBC 5.2 UNICODE Driver};Server=localhost;Database=bcbeta;User=root;Password=;Option=3"
    'On Error GoTo ErrorHandler
        conn.Open
'ErrorHandler:
 '       MsgBox "Can't Connect to Database/Server" & vbNewLine & Err.Description, vbOKOnly, "Error"
  '      End
    
End Sub
    
Public Sub dbClose()
    
    conn.Close

End Sub
'Public Function Fetch(ByVal stmt) As ADODB.Recordset
 '   Set cmd = New ADODB.Command
  '  cmd.ActiveConnection = conn
  '  cmd.CommandType = adCmdText
  '  'cmd.CommandText = "Select * from BCP_TEMP"
  '  cmd.CommandText = stmt
  '
  '  Set Fetch = cmd.Execute
'End Function

Public Function GetData() As ADODB.Recordset
    Dim stmt As String
    stmt = "Select * from bcp_temp"
     Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    'cmd.CommandText = "Select * from BCP_TEMP"
    cmd.CommandText = stmt
    
    Set GetData = cmd.Execute
    'Set GetData = Fetch(stmt)
End Function

Public Function Filter(ByVal field, ByVal sqlFrom, ByVal sqlTo, ByVal Where) As ADODB.Recordset
    Dim stmt As String
    stmt = "SELECT * FROM bcp_temp WHERE  " & Where & " `" & field & "` BETWEEN '" & sqlFrom & "' AND '" & sqlTo & "' ORDER BY Sku ASC"
    'MsgBox stmt
    
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    'cmd.CommandText = "Select * from BCP_TEMP"
    cmd.CommandText = stmt
    
    Set Filter = cmd.Execute
    'Set Filter = Fetch(stmt)
End Function

Public Function GetDept()
    Dim stmt As String
    stmt = "SELECT * FROM bcp_dept"
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = stmt
    
    Set GetDept = cmd.Execute
End Function

Public Function GetSku()
    Dim stmt As String
    stmt = "SELECT Sku FROM bcp_temp"
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = stmt
    
    Set GetSku = cmd.Execute

End Function

Public Function GetBarCode()
    Dim stmt As String
    stmt = "SELECT Barcode FROM bcp_temp"
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = stmt
    
    Set GetBarCode = cmd.Execute
End Function

Public Function GetEffectiveDate()
    Dim stmt As String
    stmt = "SELECT EffectiveDate FROM bcp_temp ORDER BY EffectiveDate ASC"
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = stmt
    
    Set GetEffectiveDate = cmd.Execute
End Function

Public Function ApplyFilters(ByVal Where)
    Dim stmt As String
    stmt = "SELECT * FROM bcp_temp WHERE " & Where
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = stmt
    
    Set ApplyFilters = cmd.Execute
End Function

Public Function insertNewBarcode(ByVal dept, ByVal sku, ByVal barcode, ByVal eDate)
    Dim stmt As String
    stmt = "INSERT into bcp_temp VALUES(NULL,'" & dept & "','" & sku & "','" & barcode & "','NA','0.00','" & eDate & "')"
    MsgBox stmt
    Set cmd = New ADODB.Command
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = stmt
    
    Set insertNewBarcode = cmd.Execute
End Function

Public Sub DBTerminate()
    Set rs = Nothing
    Set cmd = Nothing
    Set conn = Nothing
End Sub

