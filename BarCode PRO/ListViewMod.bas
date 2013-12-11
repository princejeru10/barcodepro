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
        mainMod.ResetSKU
        mainMod.ResetBarCode
        mainMod.ResetDate
        With .ListView1
            .ListItems.Clear
            rs.MoveFirst
            mainMod.SetSKUText (rs(2))
            mainMod.SetBarCodeText (rs(3))
            mainMod.SetDateText (rs(6))
            Do Until rs.EOF
                Set lstItem = .ListItems.Add(, , rs(0))
                For i = 1 To rs.Fields.Count - 1
                    lstItem.SubItems(i) = rs(i)
                Next
                mainMod.SetSKU (rs(2))
                mainMod.SetBarCode (rs(3))
                mainMod.SetDate (rs(6))
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
        Where = " DeptID BETWEEN '" & deptFrom & "' AND '" & deptTo & "' AND "
        
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

Public Function ByBarCode()
Dim rs As New ADODB.Recordset
    Dim field, bcFrom, bcTo, Where, deptFrom, deptTo As String
    
    field = "Barcode"
    
    With BCTPROmain
        bcFrom = .barCodeFrom.Text
        bcTo = .barCodeTo.Text
        deptFrom = .deptFrom.Text
        deptTo = .deptTo.Text
        Where = " DeptID BETWEEN '" & deptFrom & "' AND '" & deptTo & "' AND "
        
        DBCon.dbOpen
        Set rs = DBCon.Filter(field, bcFrom, bcTo, Where)
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
End Function

Public Function ByDate()
Dim rs As New ADODB.Recordset
    Dim field, dateFrom, dateTo, Where, deptFrom, deptTo As String
    
    field = "EffectiveDate"
    
    With BCTPROmain
        dateFrom = Format$(.dateFrom.Text, "yyyy-m-d")
        dateTo = Format$(.dateTo.Text, "yyyy-m-d")
        deptFrom = .deptFrom.Text
        deptTo = .deptTo.Text
        Where = " DeptID BETWEEN '" & deptFrom & "' AND '" & deptTo & "' AND "
        
        DBCon.dbOpen
        Set rs = DBCon.Filter(field, dateFrom, dateTo, Where)
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
End Function

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
            .DescBC1 = Mid(desc, 1, 25)
            .DescBC2 = Mid(desc, 1, 25)
            .DescBC3 = Mid(desc, 1, 25)
            .PriceBC1 = price
            .PriceBC2 = price
            .PriceBC3 = price
            .dateBC1 = eDate
            .dateBC2 = eDate
            .dateBC3 = eDate
            
    End With
End Function
