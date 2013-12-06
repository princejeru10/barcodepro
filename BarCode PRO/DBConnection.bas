Attribute VB_Name = "DBConnection"
Dim conn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim connString As String

Public Sub dbInit()

    conn.ConnectionString = "Provider=SQLNCLI10;Server=ACCTG-CHRIS\SQLEXPRESS;DataTypeCompatibility=80;Database=BCBeta;User Id=carlo_;Password=/stats;"
    conn.Open
    
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
