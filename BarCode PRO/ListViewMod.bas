Attribute VB_Name = "ListViewMod"
Dim DBCon As New DBClass

Public Sub PopulateListView()
    Dim rs As New ADODB.Recordset
    DBCon.dbOpen
        Dim lstItem As ListItem
        Set rs = DBCon.getData
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
        .skuBC2 = sku
        .skuBC3 = sku
        .skuBC4 = sku
        .skuBC5 = sku
        .skuBC6 = sku
        
        .deptBC1 = dn
        .deptBC2 = dn
        .deptBC3 = dn
        
        .BC1 = bc
        .BC2 = bc
        .BC3 = bc
        
        .BCBarCode1 = bc
        .BCBarCode2 = bc
        .BCBarCode3 = bc
        
        .DescBC1 = Mid(desc, 1, 30)
        .DescBC2 = Mid(desc, 1, 30)
        .DescBC3 = Mid(desc, 1, 30)
        
        .PriceBC1 = price
        .PriceBC2 = price
        .PriceBC3 = price
        
        .dateBC1 = eDate
        .dateBC2 = eDate
        .dateBC3 = eDate
    End With
End Function
