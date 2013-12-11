Attribute VB_Name = "PrintBarCodeMod"
Private Declare Function SendMessage Lib "user32.dll" Alias _
   "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, _
   ByVal wParam As Long, ByVal lParam As Long) As Long

Private Const WM_PAINT = &HF
Private Const WM_PRINT = &H317
Private Const PRF_CLIENT = &H4&    ' Draw the window's client area
Private Const PRF_CHILDREN = &H10& ' Draw all visible child
Private Const PRF_OWNED = &H20&    ' Draw all owned windows

Const vbHiMetric As Integer = 8

Dim PicRatio      As Double
Dim PrnWidth      As Double
Dim PrnHeight     As Double
Dim PrnRatio      As Double
Dim PrnPicWidth   As Double
Dim PrnPicHeight  As Double

Public Sub startPrint()
Dim rv As Long
    
    Printer.Print
    
    With BCTPROmain.Picture1
    'Draw controls to picture box
    rv = SendMessage(.hwnd, WM_PAINT, .hDC, 0)
    rv = SendMessage(.hwnd, WM_PRINT, .hDC, _
        PRF_CHILDREN) 'Or PRF_CLIENT Or PRF_OWNED)
    'VB.Printer.Scale (0, 0)-(200, 500)
    
    Printer.ScaleMode = 6
    'Printer.CurrentX = 3: Printer.CurrentY = 2
    .AutoRedraw = True
    .Picture = .Image
    
    ' *** Calculate device independent Width to Height
    'ratio for picture
    PicRatio = .Width / .Height
    
    
    ' *** Calculate the dimentions of the printable
    'area in HiMetric
    PrnWidth = Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbHiMetric)
    PrnHeight = Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbHiMetric)
    
    ' *** Calculate device independent Width to Height
    'ratio for printer
    PrnRatio = PrnWidth / PrnHeight
    
    ' *** Scale the output to the printable area
    If PicRatio >= PrnRatio Then
       ' *** Scale picture to fit full width of printable area
       PrnPicWidth = Printer.ScaleX(PrnWidth, vbHiMetric, _
           Printer.ScaleMode)
       PrnPicHeight = Printer.ScaleY(PrnWidth / PicRatio, _
           vbHiMetric, Printer.ScaleMode)
    Else
       ' *** Scale picture to fit full height of printable area
       PrnPicHeight = Printer.ScaleY(PrnHeight, vbHiMetric, _
           Printer.ScaleMode)
       PrnPicWidth = Printer.ScaleX(PrnHeight * PicRatio, _
           vbHiMetric, Printer.ScaleMode)
    End If
        
    'for the position of the picture on paper
    Printer.PaintPicture .Picture, 0, 0, PrnPicWidth, PrnPicHeight
    Printer.EndDoc
    .Picture = LoadPicture()
    
    End With
End Sub
