Attribute VB_Name = "mainMod"
Dim DBCon As New DBClass

Public Sub Terminate()
    DBCon.DBTerminate
End Sub

Public Sub SetDept()

    Dim rs As New ADODB.Recordset

    DBCon.dbOpen
    Set rs = DBCon.GetDept
        
    With BCTPROmain
        .deptFrom.Clear
        .deptTo.Clear
        
        .deptFrom.Text = (rs(1))
        .deptTo.Text = (rs(1))
    
        rs.MoveFirst
        
        Do Until rs.EOF
            .deptFrom.AddItem (rs(1))
            .deptTo.AddItem (rs(1))
            rs.MoveNext
        Loop
    End With
    
    DBCon.DBTerminate
    
End Sub

Public Sub ResetSKU()
    With BCTPROmain
        .skuFrom.Clear
        .skuTo.Clear
    End With
End Sub

Public Sub SetSKU(ByVal rs)
    With BCTPROmain
        .skuFrom.AddItem (rs)
        .skuTo.AddItem (rs)
        
    End With
End Sub

Public Sub SetSKUText(ByVal rs)
With BCTPROmain
        .skuFrom.Text = rs
        .skuTo.Text = rs
        
    End With
End Sub

