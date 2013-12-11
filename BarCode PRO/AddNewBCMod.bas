Attribute VB_Name = "AddNewBCMod"
Dim DBCon As New DBClass
Public Sub SetDept()

    Dim rs As New ADODB.Recordset

    DBCon.dbOpen
    Set rs = DBCon.GetDept
        
    With AddNewBarCode
        .deptText.Clear
        .deptText.Text = (rs(1))
    
        rs.MoveFirst
        
        Do Until rs.EOF
            .deptText.AddItem (rs(1))
            rs.MoveNext
        Loop
    End With
    
    DBCon.DBTerminate
    
End Sub

Public Function saveBarcode(ByVal dept, ByVal sku, ByVal barcode, ByVal eDate)
    DBCon.dbOpen
    Set saveBarcode = DBCon.insertNewBarcode(dept, sku, barcode, eDate)
    DBCon.DBTerminate
End Function

Public Function getCN() As Integer
    Dim deptNum As String
    Dim sku As String
    Dim barcode As String
    Dim ctr As Integer 'Counter
    Dim p1 As Integer 'even
    Dim p2 As Integer 'odd
    Dim cn As Integer
    Dim z, r As Integer
    Dim temp As Integer
    
    ctr = 0
    p1 = 0
    p2 = 0
    With AddNewBarCode
        deptNum = .deptText.Text
        sku = .skuText.Text
        barcode = deptNum & sku
    End With
    Do Until ctr >= Len(barcode)
        temp = Mid(barcode, ctr + 1, 1)
        If Not (ctr Mod 2) = 0 Then
            p1 = p1 + (temp * 3)
        Else
            p2 = p2 + temp
        End If
        ctr = ctr + 1
    Loop
        
    r = 0
    z = p1 + p2
    
    r = NearestTen(z, r)
    
        
    cn = r - z
    
    getCN = cn
End Function

Public Function NearestTen(ByVal z As Integer, ByRef r As Integer) As Integer
    Dim temp As Integer
    
    If (z Mod 10) > 0 Then
        temp = 10 - (z Mod 10)
        r = z + temp
    End If
    
    NearestTen = r
End Function
