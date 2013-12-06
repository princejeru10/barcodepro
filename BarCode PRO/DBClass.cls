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
Dim connString As String

Public Sub dbOpen()

    conn.ConnectionString = "Provider=SQLNCLI10;Server=ACCTG-CHRIS\SQLEXPRESS;DataTypeCompatibility=80;Database=BCBeta;User Id=carlo_;Password=/stats;"
    On Error GoTo ErrorHandler
        conn.Open
ErrorHandler:
        MsgBox "Can't Connect to Database/Server", vbOKOnly, "Error"
        End
    
End Sub
    
Public Sub dbClose()
    
    conn.Close

End Sub

Public Function getData() As ADODB.Recordset
    
    rs.Open "Select * from BCP_TEMP", conn, adOpenDynamic
        
End Function

Public Sub rsClose()
    rs.Close
End Sub

