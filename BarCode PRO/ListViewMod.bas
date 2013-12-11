Attribute VB_Name = "ListViewMod"
Dim DBCon As New DBClass

Public Sub PopulateListView()
    Dim rs As New ADODB.Recordset
    DBCon.dbOpen
        Dim lstItem As ListItem
        Set rs = DBCon.GetData
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


Public Sub ByDept()
    Dim rs As New ADODB.Recordset
    Dim field, deptFrom, deptTo, Where As String
    
    field = "DeptID"
    
    With BCTPROmain
        deptFrom = .deptFrom.Text
        deptTo = .deptTo.Text
        .skuBtn.Enabled = True
        Where = ""
        
        DBCon.dbOpen
        Set rs = DBCon.Filter(field, deptFrom, deptTo, Where)
    
        With .ListView1
            .ListItems.Clear
            rs.MoveFirst
            mainMod.SetSKUText (rs(2))
            Do Until rs.EOF
                Set lstItem = .ListItems.Add(, , rs(0))
                For i = 1 To rs.Fields.Count - 1
                    lstItem.SubItems(i) = rs(i)
                Next
                mainMod.SetSKU (rs(2))
            rs.MoveNext
            Loop
        End With
    End With
    DBCon.DBTerminate
            
    SelectFirstRow
End Sub

Public Sub BySKU()
    Dim rs As New ADODB.Recordset
    Dim field, skuFrom, skuTo, Where, deptFrom, deptTo As String
    
    field = "Sku"
    
    With BCTPROmain
        skuFrom = .skuFrom.Text
        skuTo = .skuTo.Text
        deptFrom = .deptFrom.Text
        deptTo = .deptTo.Text
        Where = "AND WHERE DeptID BETWEEN '" & deptFrom & "' AND '" & deptTo & "'"
        
        DBCon.dbOpen
        Set rs = DBCon.Filter(field, skuFrom, skuTo, Where)
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

Public Function SelectFirstRow()
    With BCTPROmain
        With .ListView1
            .ListItems(1).Selected = True
        End With
    End With
    
    ChangeBCValues
End Function

Public Function ChangeBCValues()
    With BCTPROmain
        Dim dn, sku, bc, desc, price, eDate As String
        With .ListView1.SelectedItem
            dn = .SubItems(1)
            sku = .SubItems(2)
            bc = .SubItems(3)
            desc = .SubItems(4)
            price = .SubItems(5)
            eDate = .SubItems(6)
        End With
        
        .skuBC1 = sku
        .skuBC4 = sku
        
        .deptBC1 = dn
        
        .BC1 = bc
        
        .BCBarCode1 = bc
        
        .DescBC1 = Mid(desc, 1, 30)
        
        .PriceBC1 = price
        
        .dateBC1 = eDate
    End With
End Function
