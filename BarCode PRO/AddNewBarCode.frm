VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form AddNewBarCode 
   Caption         =   "Add New Barcode"
   ClientHeight    =   2820
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4125
   LinkTopic       =   "Form1"
   ScaleHeight     =   2820
   ScaleWidth      =   4125
   StartUpPosition =   3  'Windows Default
   Begin MSComCtl2.DTPicker eDateText 
      Height          =   375
      Left            =   1440
      TabIndex        =   8
      Top             =   1560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      _Version        =   393216
      Format          =   40960001
      CurrentDate     =   41619
   End
   Begin VB.CommandButton saveBtn 
      Caption         =   "Save BarCode"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   2040
      Width           =   2175
   End
   Begin VB.TextBox barcodeText 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1200
      Width           =   2175
   End
   Begin VB.TextBox skuText 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   720
      Width           =   1935
   End
   Begin VB.ComboBox deptText 
      Height          =   315
      Left            =   1440
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Effective Date"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Barcode"
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "SKU"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Department ID"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "AddNewBarCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    AddNewBCMod.SetDept
    resetText
End Sub

Private Sub resetText()
    skuText.Text = ""
    barcodeText.Text = ""
End Sub

Private Sub saveBtn_Click()
    Dim dept, sku, barcode, eDate As String
    Dim fn
    
    dept = deptText.Text
    sku = skuText.Text
    barcode = barcodeText.Text
    'eDate = year.Text & "-" & month.Text & "-" & day.Text
    eDate = Format$(eDateText.Value, "yyyy-m-d")
    'MsgBox eDate
    fn = AddNewBCMod.saveBarcode(dept, sku, barcode, eDate)
End Sub

Private Sub skuText_LostFocus()
    Dim bc As String
    barcodeText.Text = deptText.Text & skuText.Text & AddNewBCMod.getCN
    bc = barcodeText.Text
    If (Len(bc) < 13) Then
    MsgBox ("Please check correct values for barcode! Barcode length should be 13 digits!" & Len(bc))
    saveBtn.Enabled = False
    Else
    saveBtn.Enabled = True
    End If
End Sub
