Attribute VB_Name = "mainMod"
Dim DBCon As New DBClass

Public Sub Terminate()
    DBCon.DBTerminate
End Sub

Public Sub SetDept()
    Dim rs As New ADODB.Recordset
    Set rs = DBCon.GetDept
    
    BCTPROmain.Combo1.Text = rs(1)
    
End Sub
