VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Queries"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim DBCon As New DBClass

Public Function Filter(ByVal field, ByVal sqlFrom, ByVal sqlTo) As ADODB.Recordset
    Dim stmt As String
    stmt = "SELECT * FROM bcp_temp WHERE " & field & " BETWEEN " & sqlFrom & " AND " & sqlTo
    DBCon.fetch (stmt)
End Function
