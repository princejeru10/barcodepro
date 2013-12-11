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
        .skuBtn.Enabled = True
    End With
End Sub

Public Sub SetSKU(ByVal rs As String)
    'Dim rs As New ADODB.Recordset

    'DBCon.dbOpen
    
    'Set rs = DBCon.GetSku
        
    With BCTPROmain
        'rs.MoveFirst
        
        'Do Until rs.EOF
            .skuFrom.AddItem (rs)
            .skuTo.AddItem (rs)
         '   rs.MoveNext
        'Loop
        
    End With
    
    'DBCon.DBTerminate
End Sub

Public Sub SetSKUText(ByVal rs)
    With BCTPROmain
        .skuFrom.Text = rs
        .skuTo.Text = rs
        
    End With
End Sub


Public Sub SetBarCode()
    
    Dim rs As New ADODB.Recordset

    DBCon.dbOpen
    
    Set rs = DBCon.GetBarCode
        
    With BCTPROmain
        .barCodeFrom.Clear
        .barCodeTo.Clear
        
        .barCodeFrom.Text = rs(0)
        .barCodeTo.Text = rs(0)
        
        rs.MoveFirst
        
        Do Until rs.EOF
            .barCodeFrom.AddItem (rs(0))
            .barCodeTo.AddItem (rs(0))
            rs.MoveNext
        Loop
        
    End With
    
    DBCon.DBTerminate
End Sub

Public Sub SetEffectiveDate()
    Dim rs As New ADODB.Recordset

    DBCon.dbOpen
    
    Set rs = DBCon.GetEffectiveDate
        
    With BCTPROmain
        .dateFrom.Clear
        .dateTo.Clear
        
        .dateFrom.Text = rs(0)
        .dateTo.Text = rs(0)
        
        rs.MoveFirst
        
        Do Until rs.EOF
            .dateFrom.AddItem (rs(0))
            .dateTo.AddItem (rs(0))
            rs.MoveNext
        Loop
        
    End With
    
    DBCon.DBTerminate
End Sub

Public Sub SetFilteredValues()
    
    Dim rs As New ADODB.Recordset
    Dim Where As String
    DBCon.dbOpen
        Dim lstItem As ListItem
        
        With BCTPROmain
            Where = "DeptID BETWEEN '" & .deptFrom.Text & "' AND '" & .deptTo.Text & "' AND Sku BETWEEN '" & .skuFrom.Text & "' AND '" & .skuTo.Text & "'"
            
        End With
        Set rs = DBCon.ApplyFilters(Where)
            With BCTPROmain
                With .ListView1
                    .ListItems.Clear
                    rs.MoveFirst
                    Do Until rs.EOF
                        Set lstItem = .ListItems.Add(, , rs(0))
                        For i = 1 To rs.Fields.Count - 1
                            lstItem.SubItems(i) = rs(i)
                        Next
                    rs.MoveNext
                    Loop
                End With
            End With
    DBCon.DBTerminate
            
    SelectFirstRow
End Sub
