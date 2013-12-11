VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
   Begin VB.CommandButton ClearBtn 
      Caption         =   "RESET"
      Height          =   495
      Left            =   120
      TabIndex        =   26
      Top             =   6960
      Width           =   2295
   End
   Begin VB.CommandButton submitBtn 
      Caption         =   "SUBMIT"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2520
      TabIndex        =   21
      Top             =   6960
      Width           =   2295
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3375
      Left            =   4920
      TabIndex        =   14
      Top             =   3360
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   5953
      View            =   3
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "ID"
         Object.Width           =   882
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Department ID"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "SKU"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Barcode"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Description"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Price"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Effective Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton ExitBtn 
      Caption         =   "EXIT"
      Height          =   495
      Left            =   13560
      TabIndex        =   13
      Top             =   6840
      Width           =   2055
   End
   Begin VB.CommandButton PrintBtn 
      Caption         =   "PRINT"
      Height          =   495
      Left            =   11280
      TabIndex        =   12
      Top             =   6840
      Width           =   2175
   End
   Begin VB.Frame dateFrame 
      Caption         =   "EFFECTIVE DATE"
      Height          =   1575
      Left            =   120
      TabIndex        =   9
      Top             =   5160
      Width           =   4695
      Begin VB.CommandButton DateBtn 
         Caption         =   "SUBMIT"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2880
         TabIndex        =   28
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox dateTo 
         Height          =   315
         Left            =   2880
         TabIndex        =   25
         Text            =   "Combo2"
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox dateFrom 
         Height          =   315
         Left            =   720
         TabIndex        =   24
         Text            =   "Combo1"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label dateFromLabel 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   480
         Width           =   735
      End
      Begin VB.Label dateToLabel 
         Caption         =   "To"
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   480
         Width           =   255
      End
   End
   Begin VB.Frame barcodeFrame 
      Caption         =   "BARCODE/ITEM"
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   4695
      Begin VB.CommandButton barcodeBtn 
         Caption         =   "SUBMIT"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2880
         TabIndex        =   27
         Top             =   960
         Width           =   1575
      End
      Begin VB.ComboBox barCodeTo 
         Height          =   315
         Left            =   2880
         TabIndex        =   23
         Text            =   "Combo1"
         Top             =   480
         Width           =   1575
      End
      Begin VB.ComboBox barCodeFrom 
         Height          =   315
         Left            =   720
         TabIndex        =   22
         Text            =   "Combo1"
         Top             =   480
         Width           =   1575
      End
      Begin VB.Label barcodeFromLabel 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   480
         Width           =   735
      End
      Begin VB.Label barcodeToLabel 
         Caption         =   "To"
         Height          =   375
         Left            =   2520
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame skuFrame 
      Caption         =   "SKU"
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   1800
      Width           =   4695
      Begin VB.ComboBox skuTo 
         Height          =   315
         Left            =   2880
         TabIndex        =   20
         Text            =   "Combo1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox skuFrom 
         Height          =   315
         Left            =   720
         TabIndex        =   19
         Text            =   "Combo1"
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton skuBtn 
         Caption         =   "SUBMIT"
         Enabled         =   0   'False
         Height          =   495
         Left            =   2880
         TabIndex        =   18
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label skuFromLabel 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Width           =   495
      End
      Begin VB.Label skuToLabel 
         Caption         =   "To"
         Height          =   375
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   255
      End
   End
   Begin VB.Frame deptFrame 
      Caption         =   "DEPARTMENT"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.CommandButton deptFilterBtn 
         Caption         =   "SUBMIT"
         Height          =   495
         Left            =   2880
         TabIndex        =   17
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox deptTo 
         Height          =   315
         Left            =   2880
         TabIndex        =   16
         Text            =   "Combo2"
         Top             =   360
         Width           =   1575
      End
      Begin VB.ComboBox deptFrom 
         Height          =   315
         Left            =   720
         TabIndex        =   15
         Text            =   "Combo1"
         Top             =   360
         Width           =   1590
      End
      Begin VB.Label deptToLabel 
         Caption         =   "To"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         Top             =   360
         Width           =   375
      End
      Begin VB.Label deptFromLabel 
         Caption         =   "From"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   3015
      Left            =   4920
      ScaleHeight     =   2955
      ScaleWidth      =   10635
      TabIndex        =   29
      Top             =   240
      Width           =   10695
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12/5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   10080
         TabIndex        =   62
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12345678"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   61
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1234"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   60
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890128"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   59
         Top             =   2280
         Width           =   3015
      End
      Begin VB.Label Label21 
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
         Left            =   7320
         TabIndex        =   58
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM DESCRIPTION"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7680
         TabIndex        =   57
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12345678"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   56
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "/PC"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9840
         TabIndex        =   55
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7800
         TabIndex        =   54
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9999.99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7320
         TabIndex        =   53
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FISHER FASHION STORE"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium Cond"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   7440
         TabIndex        =   52
         Top             =   0
         Width           =   3135
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   7440
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12/5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         TabIndex        =   51
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12345678"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   49
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890128"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   48
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label10 
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
         Left            =   3720
         TabIndex        =   47
         Top             =   1560
         Width           =   3375
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM DESCRIPTION"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4080
         TabIndex        =   46
         Top             =   1200
         Width           =   2655
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12345678"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   45
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "/PC"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   44
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   43
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9999.99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3720
         TabIndex        =   42
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FISHER FASHION STORE"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium Cond"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   3720
         TabIndex        =   41
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label BCBarCode1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890128"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1200
         TabIndex        =   40
         Top             =   2280
         Width           =   1335
      End
      Begin VB.Label BC1 
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
         Left            =   120
         TabIndex        =   39
         Top             =   1560
         Width           =   3255
      End
      Begin VB.Label dateBC1 
         BackStyle       =   0  'Transparent
         Caption         =   "12/5"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2640
         TabIndex        =   38
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label deptBC1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1234"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   600
         TabIndex        =   37
         Top             =   2520
         Width           =   2415
      End
      Begin VB.Label skuBC4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "12345678"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   480
         TabIndex        =   36
         Top             =   2520
         Width           =   855
      End
      Begin VB.Label DescBC1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM DESCRIPTION"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   35
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label skuBC1 
         Alignment       =   2  'Center
         Caption         =   "12345678"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "/PC"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   33
         Top             =   600
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "P"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   32
         Top             =   600
         Width           =   495
      End
      Begin VB.Label PriceBC1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "9999.99"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   31
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "FISHER FASHION STORE"
         BeginProperty Font 
            Name            =   "Franklin Gothic Medium Cond"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000B&
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   0
         Width           =   3255
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   240
         Top             =   0
         Width           =   3015
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   3840
         Top             =   0
         Width           =   3135
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "1234"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3840
         TabIndex        =   50
         Top             =   2520
         Width           =   3135
      End
   End
End
Attribute VB_Name = "BCTPROmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Public bc As String


Private Sub ClearBtn_Click()
    reset
    ListViewMod.PopulateListView
End Sub

Private Sub deptFilterBtn_Click()
    If deptFrom.Text > deptTo.Text Then
    MsgBox "Departmet From must be lower than Department To!", , "Please check correct values"
    Else
    ListViewMod.ByDept
    End If
End Sub

Private Sub PrintBtn_Click()
Dim item As ListItem
    For Each item In ListView1.ListItems
        If item.Selected = True Then
            With item
            dn = .SubItems(1)
            sku = .SubItems(2)
            bc = .SubItems(3)
            desc = .SubItems(4)
            price = .SubItems(5)
            eDate = .SubItems(6)
            End With
            
            skuBC1 = sku
            skuBC4 = sku
            
            deptBC1 = dn
            
            BC1 = bc
            
            BCBarCode1 = bc
            
            DescBC1 = Mid(desc, 1, 30)
            
            PriceBC1 = price
            
            dateBC1 = eDate
            
            PrintBarCodeMod.startPrint
        End If
    Next
End Sub

Private Sub skuBtn_Click()
    ListViewMod.BySKU
End Sub


Private Sub Form_Load()
    reset
    ListViewMod.PopulateListView
    mainMod.SetDept
    'mainMod.SetSKU
    'mainMod.SetBarCode
    'mainMod.SetEffectiveDate
    
End Sub

Private Function reset()
    'deptFrom.Text = ""
    'deptTo.Text = ""
    'skuFrom.Text = ""
    skuFrom.Clear
    skuTo.Clear 'Text = ""
    barCodeFrom.Clear '.Text = ""
    barCodeTo.Clear 'Text = ""
    dateFrom.Clear 'Text = ""
    dateTo.Clear 'Text = ""
    
    skuBtn.Enabled = False
    barcodeBtn.Enabled = False
    DateBtn.Enabled = False
End Function

Private Sub ListView1_Click()
    ListViewMod.ChangeBCValues
End Sub

Private Sub PreviewBtn_Click()
    Dim cn As Integer
    Dim deptNum, sku As String
    
    deptNum = deptFromText.Text
    sku = skuFromText.Text
    
End Sub

Private Sub resetBtn_Click()
    reset
End Sub

Private Sub PrintBarCode()
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

Private Function NearestTen(ByVal z As Integer, ByRef r As Integer) As Integer
    Dim temp As Integer
    
    If (z Mod 10) > 0 Then
        temp = 10 - (z Mod 10)
        r = z + temp
    End If
    
    NearestTen = r
End Function

Private Sub ExitBtn_Click()
    mainMod.Terminate
    End
End Sub

Private Sub submitBtn_Click()
    mainMod.SetFilteredValues
End Sub
