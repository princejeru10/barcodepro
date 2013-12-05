VERSION 5.00
Begin VB.MDIForm BarCodePRO 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   8625
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   14715
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu File 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu New 
         Caption         =   "&New"
      End
      Begin VB.Menu Exit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "BarCodePRO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub New_Click()
    Filter.Show
End Sub
