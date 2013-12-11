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

Public Sub ResetBarCode()
    With BCTPROmain
        .barCodeFrom.Clear
        .barCodeTo.Clear
        .barcodeBtn.Enabled = True
        End With
End Sub

Public Sub SetBarCodeText(ByVal rs) 'set barcode/item drop down text
    With BCTPROmain
        .barCodeFrom.Text = rs
        .barCodeTo.Text = rs
    End With
End Sub

Public Sub SetBarCode(ByVal rs As String)
            
    With BCTPROmain
        
    .barCodeFrom.AddItem (rs)
    .barCodeTo.AddItem (rs)
        
    End With
End Sub

Public Sub ResetDate()
    With BCTPROmain
        .dateFrom.Clear
        .dateTo.Clear
        .DateBtn.Enabled = True
        End With
End Sub

Public Sub SetDateText(ByVal rs) 'set barcode/item drop down text
    With BCTPROmain
        .dateFrom.Text = rs
        .dateTo.Text = rs
    End With
End Sub

Public Sub SetDate(ByVal rs As String)
            
    With BCTPROmain
        If .dateFrom.ListCount = 0 Then GoTo doAadd
        
        Dim i As Integer
        For i = 0 To .dateFrom.ListCount - 1
            If .dateFrom.List(i) <> rs Then
                .dateFrom.AddItem (rs)
                .dateTo.AddItem (rs)
                Exit Sub
            End If
        Next
        
doAdd:
        .dateFrom.AddItem (rs)
        .dateTo.AddItem (rs)
        
    End With
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
