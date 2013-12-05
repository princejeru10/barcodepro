VERSION 5.00
Object = "{ED7AED90-6B75-406A-A856-CEFDC6E021BB}#1.243#0"; "STROKE~1.OCX"
Begin VB.Form BCTPROmain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "BarCode Tender PRO Version Beta"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   15735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   15735
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton PrintAllBtn 
      Caption         =   "PRINT ALL"
      Height          =   495
      Left            =   11400
      TabIndex        =   56
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton ExitBtn 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   13560
      TabIndex        =   55
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton PrintBtn 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   9240
      TabIndex        =   54
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton NextBtn 
      Caption         =   "NEXT"
      Height          =   495
      Left            =   7080
      TabIndex        =   53
      Top             =   3360
      Width           =   2055
   End
   Begin VB.CommandButton PrevBtn 
      Caption         =   "PREVIOUS"
      Height          =   495
      Left            =   4920
      TabIndex        =   52
      Top             =   3360
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   4920
      ScaleHeight     =   2955
      ScaleWidth      =   10635
      TabIndex        =   22
      Top             =   240
      Width           =   10695
      Begin STROKESCRIBELibCtl.StrokeScribe StrokeScribe1 
         Height          =   975
         Index           =   0
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Width           =   3255
         _Version        =   65779
         FontColor       =   0
         BkgndColor      =   16777215
         TextBelow       =   ""
         _ExtentX        =   5741
         _ExtentY        =   1720
         _StockProps     =   64
      End
      Begin STROKESCRIBELibCtl.StrokeScribe StrokeScribe1 
         Height          =   975
         Index           =   1
         Left            =   3600
         TabIndex        =   24
         Top             =   1560
         Width           =   3255
         _Version        =   65779
         FontColor       =   0
         BkgndColor      =   16777215
         TextBelow       =   ""
         _ExtentX        =   5741
         _ExtentY        =   1720
         _StockProps     =   64
      End
      Begin VB.Label BCBarCode3 
         Alignment       =   2  'Center
         Caption         =   "Label17"
         Height          =   255
         Left            =   7080
         TabIndex        =   59
         Top             =   2280
         Width           =   3255
      End
      Begin VB.Label BC3 
         Alignment       =   2  'Center
         Caption         =   "1234567890128"
         BeginProperty Font 
            Name            =   "EAN-13"
            Size            =   72
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   7200
         TabIndex        =   58
         Top             =   1560
         Width           =   3015
      End
      Begin VB.Label dateBC3 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         Height          =   255
         Left            =   9360
         TabIndex        =   51
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label deptBC3 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         Height          =   255
         Left            =   8280
         TabIndex        =   50
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label skuBC6 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         Height          =   255
         Left            =   7200
         TabIndex        =   49
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label dateBC2 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         Height          =   255
         Left            =   5880
         TabIndex        =   48
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label deptBC2 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         Height          =   255
         Left            =   4800
         TabIndex        =   47
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label skuBC5 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         Height          =   255
         Left            =   3720
         TabIndex        =   46
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label dateBC1 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         Height          =   255
         Left            =   2400
         TabIndex        =   45
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label deptBC1 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         Height          =   255
         Left            =   1320
         TabIndex        =   44
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label skuBC4 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         Caption         =   "G-BARE-SATINSHORTS/12/OYSTER"
         Height          =   255
         Left            =   7080
         TabIndex        =   42
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Caption         =   "G-BARE-SATINSHORTS/12/OYSTER"
         Height          =   255
         Left            =   3600
         TabIndex        =   41
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "G-BARE-SATINSHORTS/12/OYSTER"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   1320
         Width           =   3255
      End
      Begin VB.Label skuBC3 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         Height          =   255
         Left            =   7080
         TabIndex        =   39
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label skuBC2 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         Height          =   255
         Left            =   3600
         TabIndex        =   38
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label skuBC1 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   1080
         Width           =   3255
      End
      Begin VB.Label Label12 
         Caption         =   "/PC"
         Height          =   255
         Left            =   9480
         TabIndex        =   36
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7680
         TabIndex        =   35
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "99.99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7080
         TabIndex        =   34
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label9 
         Caption         =   "/PC"
         Height          =   255
         Left            =   6000
         TabIndex        =   33
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         TabIndex        =   32
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "999.99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3600
         TabIndex        =   31
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label6 
         Caption         =   "/PC"
         Height          =   255
         Left            =   2520
         TabIndex        =   30
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         TabIndex        =   29
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "9999.99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   28
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FISHER FASHION"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium Cond"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   7080
         TabIndex        =   27
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FISHER FASHION"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium Cond"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   735
         Left            =   -120
         TabIndex        =   26
         Top             =   120
         Width           =   10695
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FISHER FASHION"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium Cond"
            Size            =   21.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   495
         Left            =   120
         TabIndex        =   25
         Top             =   120
         Width           =   3255
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   7080
         Top             =   120
         Width           =   3255
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   3600
         Top             =   120
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   120
         Top             =   120
         Width           =   3255
      End
   End
   Begin VB.CommandButton ResetBtn 
      Caption         =   "RESET"
      Height          =   495
      Left            =   2520
      TabIndex        =   21
      Top             =   6840
      Width           =   2295
   End
   Begin VB.CommandButton PreviewBtn 
      Caption         =   "PREVIEW"
      Default         =   -1  'True
      Height          =   495
      Left            =   120
      TabIndex        =   20
      Top             =   6840
      Width           =   2295
   End
   Begin VB.Frame dateFrame 
      Caption         =   "EFFECTIVE DATE"
      Height          =   1575
      Left            =   120
      TabIndex        =   15
      Top             =   5160
      Width           =   4695
      Begin VB.TextBox dateFromText 
         Height          =   405
         Left            =   840
         TabIndex        =   16
         Text            =   "Text1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox dateToText 
         Height          =   405
         Left            =   840
         TabIndex        =   17
         Text            =   "Text1"
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label dateFromLabel 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   480
         Width           =   735
      End
      Begin VB.Label dateToLabel 
         Caption         =   "To"
         Height          =   375
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame barcodeFrame 
      Caption         =   "BARCODE/ITEM"
      Height          =   1575
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   4695
      Begin VB.TextBox barcodeFromText 
         Height          =   405
         Left            =   840
         MaxLength       =   13
         TabIndex        =   11
         Text            =   "Text1"
         Top             =   360
         Width           =   3615
      End
      Begin VB.TextBox barcodeToText 
         Height          =   405
         Left            =   840
         MaxLength       =   13
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label barcodeFromLabel 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   735
      End
      Begin VB.Label barcodeToLabel 
         Caption         =   "To"
         Height          =   375
         Left            =   240
         TabIndex        =   13
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame skuFrame 
      Caption         =   "SKU"
      Height          =   1575
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   4695
      Begin VB.TextBox skuFromText 
         Height          =   405
         Left            =   840
         MaxLength       =   8
         TabIndex        =   6
         Text            =   "Text1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.TextBox skuToText 
         Height          =   405
         Left            =   840
         MaxLength       =   8
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label skuFromLabel 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   735
      End
      Begin VB.Label skuToLabel 
         Caption         =   "To"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Frame deptFrame 
      Caption         =   "DEPARTMENT"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.TextBox deptToText 
         Height          =   405
         Left            =   840
         MaxLength       =   4
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox deptFromText 
         Height          =   405
         Left            =   840
         MaxLength       =   4
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label deptToLabel 
         Caption         =   "To"
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
      Begin VB.Label deptFromLabel 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Label Label16 
      Caption         =   "1234567890128"
      BeginProperty Font 
         Name            =   "EAN-13"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5280
      TabIndex        =   57
      Top             =   1800
      Width           =   2535
   End
End
Attribute VB_Name = "BCTPROmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public bc As String

Private Sub Form_Load()
    reset
End Sub

Public Function reset()
    deptFromText.Text = ""
    deptToText.Text = ""
    skuFromText.Text = ""
    skuToText.Text = ""
    barcodeFromText.Text = ""
    barcodeToText.Text = ""
    dateFromText.Text = ""
    dateToText.Text = ""
End Function



Private Sub PreviewBtn_Click()
    Dim cn As Integer
    Dim deptNum, sku As String
    
    deptNum = deptFromText.Text
    sku = skuFromText.Text
    
    cn = getCN
    bc = deptNum & sku & cn
    
    skuBC1.Caption = sku
    skuBC2.Caption = sku
    skuBC3.Caption = sku
    skuBC4.Caption = sku
    skuBC5.Caption = sku
    skuBC6.Caption = sku
    
    deptBC1.Caption = deptNum
    deptBC2.Caption = deptNum
    deptBC3.Caption = deptNum
    
    dateBC1.Caption = Month(Now) & "/" & Day(Now)
    dateBC2.Caption = Month(Now) & "/" & Day(Now)
    dateBC3.Caption = Month(Now) & "/" & Day(Now)
    
    BCbarCode3.Caption = bc
    
    StrokeScribe1(0).Text = deptNum & sku
    StrokeScribe1(1).Text = deptNum & sku
    BC3.Caption = bc
End Sub

Private Sub resetBtn_Click()
    reset
End Sub

Private Sub printBarCode()
    Printer.PaintPicture Picture1.Picture, 25, 25
    Printer.EndDoc
End Sub

Private Function getCN() As Integer
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
    
    deptNum = deptFromText.Text
    sku = skuFromText.Text
    barcode = deptNum & sku
        
    Do Until ctr >= Len(barcode)
        temp = Mid(barcode, ctr + 1, 1)
        If Not (ctr Mod 2) = 0 Then
            p1 = p1 + temp
        Else
            p2 = p2 + temp
        End If
        ctr = ctr + 1
    Loop
        
    r = 0
    z = p1 + 3 * p2
    
    r = NearestTen(z, r)
        
    cn = r - z
    
    getCN = cn
End Function

Private Function NearestTen(ByVal z As Integer, ByRef r As Integer) As Integer
    Dim temp As Integer
    
    If (z Mod 10) > 0 Then
        temp = 10 - (z Mod 10)
        r = z + temp
    End If
    
    NearestTen = r
End Function

Private Sub ExitBtn_Click()
    End
End Sub

