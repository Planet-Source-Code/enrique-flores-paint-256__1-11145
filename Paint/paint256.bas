Attribute VB_Name = "Module2"
Option Explicit

' ------- my constants -----
' Draw Actions ~ instruments
Public Const A_SELECT = 0
Public Const A_PEN = 1
Public Const A_LINE = 2
Public Const A_RRECT = 3
Public Const A_FRRECT = 4
Public Const A_FILL = 5
Public Const A_FORMULA = 6
Public Const A_TEXT = 7
Public Const A_BEVEL = 8
Public Const A_ARROW = 9

' Fill Types
Public Const VT_NONE = 0         ' VT_ from dutch: Vul Type
Public Const VT_SOLID = 1
Public Const VT_FLOOD = 2
Public Const VT_FLOODRAS = 3
Public Const VT_CLIPBRAS = 4

' Flood Types
Public Const FT_LEFTRIGHT = 0
Public Const FT_LEFTRIGHT2 = 1
Public Const FT_UPDOWN = 2
Public Const FT_UPDOWN2 = 3
Public Const FT_ULDR = 4         ' upper-left to bottom-right, etc.
Public Const FT_ULDR2 = 5
Public Const FT_DLUR = 6
Public Const FT_DLUR2 = 7
Public Const FT_CIRCLE = 8
Public Const FT_SQUARE = 9

'---------- my variables ------
' Color Pallet
Public NUseColor(16) As Long     ' When setting a new
Public NColorSet(256) As Long    ' pallet
Public NBaseColcnt As Long
Public NMakeFluent As Long

Public UseColor(16) As Long      ' base colors
Public ColorSet(256) As Long     ' pallet colors
Public BaseColcnt As Long        ' base color count
Public MakeFluent As Long        '
Public BGCol As Long             ' background color

Public SzeX As Long, SzeY As Long ' size All
Public PSzeX As Long, PSzeY As Long ' size Pattern pic
Public Const MaxSize = 1024
Public PixSze As Long            ' size 1 pixel after scaling
Public ShowGrid As Boolean       ' show grid yes or no

Public OK As Boolean             ' to check after dialog

Public Const PI = 3.1415928

Public ClipMode As Long          ' insert image (open or paste)
Public AdaptTo As Long

Public FillType As Long          ' normal, fluent, fluent box, clipb. box
Public FloodType As Long         ' which pattern used for flood

Public Shade As Long             ' Shade type 0:none or 1-9 directions
Public ShadeDis As Long          '         distance
Public ShadeCol As Long          '         color 0:white, 1:black, 2:left 3:right
Public ShadeDouble As Long       ' opposite direction too? 0-1

Public FormulaPts As Long        ' points
Public FormulaGplus As Long      ' Angle = Angle + GPlus ~ Angle step
Public FormulaAngle As Long      ' start from Angle
Public FormulaFilled As Long

Public ArrowAngle As Long        ' up-down, left-right
Public ArrowWidth As Long        ' line width
Public ArrowHeadW As Long        ' head width
Public ArrowDouble As Boolean    ' two heads
Public ArrowFilled As Boolean
Public ArrowIndex As Long

Public PenWidth As Long
Public Rounding As Long          ' rectangle to circle %

Public Text As String
Public TextBold As Boolean
Public TextItalic As Boolean
Public TextSize As Single
Public TextFntNm As String
Public TextAngle As Long         ' up-down, left-right
Public TextAlign  As Long
Public TextdX As Long, TextdY As Long

' for frmBMP
Public DrivePaths() As String    ' enabling to restore prev. drive when selected one generated an error
Public prevDrive As String
Public prevPath As String

'--------- API's ------------
Type POINTAPI
        X As Long
        Y As Long
End Type

Declare Function Polyline Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Declare Function PtVisible Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Declare Function Rectangle Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long
Declare Function Pie Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function Arc Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function Chord Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long, ByVal X4 As Long, ByVal Y4 As Long) As Long
Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
Declare Function FloodFill Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal crColor As Long) As Long
Declare Function SetPolyFillMode Lib "gdi32" (ByVal hdc As Long, ByVal nPolyFillMode As Long) As Long
Public Const FLOODFILLBORDER = 0
Public Const FLOODFILLSURFACE = 1

Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long

Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Public Const SRCPAINT = &HEE0086    ' (DWORD) dest = source OR dest
Public Const SRCAND = &H8800C6  ' (DWORD) dest = source AND dest
Public Const SRCINVERT = &H660046 ' (DWORD) dest = source XOR dest

Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type
Public RC As RECT

Type Size
        cX As Long
        cY As Long
End Type

Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, lpDx As Long) As Long
Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetTextExtentPoint Lib "gdi32" Alias "GetTextExtentPointA" (ByVal hdc As Long, ByVal lpszString As String, ByVal cbString As Long, lpSize As Size) As Long

Public Const DT_TOP = &H0
Public Const DT_LEFT = &H0
Public Const DT_BOTTOM = &H8
Public Const DT_CALCRECT = &H400
Public Const DT_CENTER = &H1
Public Const DT_EXPANDTABS = &H40
Public Const DT_EXTERNALLEADING = &H200
Public Const DT_INTERNAL = &H1000
Public Const DT_NOCLIP = &H100
Public Const DT_NOPREFIX = &H800
Public Const DT_RIGHT = &H2
Public Const DT_SINGLELINE = &H20
Public Const DT_TABSTOP = &H80
Public Const DT_VCENTER = &H4
Public Const DT_WORDBREAK = &H10

''''''''
Type TEXTMETRIC
        tmHeight As Long
        tmAscent As Long
        tmDescent As Long
        tmInternalLeading As Long
        tmExternalLeading As Long
        tmAveCharWidth As Long
        tmMaxCharWidth As Long
        tmWeight As Long
        tmOverhang As Long
        tmDigitizedAspectX As Long
        tmDigitizedAspectY As Long
        tmFirstChar As Byte
        tmLastChar As Byte
        tmDefaultChar As Byte
        tmBreakChar As Byte
        tmItalic As Byte
        tmUnderlined As Byte
        tmStruckOut As Byte
        tmPitchAndFamily As Byte
        tmCharSet As Byte
End Type

Public TM As TEXTMETRIC
''''''''
Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Long, ByVal W As Long, ByVal E As Long, ByVal O As Long, ByVal W As Long, ByVal I As Long, ByVal U As Long, ByVal S As Long, ByVal C As Long, ByVal OP As Long, ByVal CP As Long, ByVal Q As Long, ByVal PAF As Long, ByVal F As String) As Long
Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Long, lpMetrics As TEXTMETRIC) As Long
Declare Function GetTextFace Lib "gdi32" Alias "GetTextFaceA" (ByVal hdc As Long, ByVal nCount As Long, ByVal lpFacename As String) As Long

'BITMAP

Type BITMAP
        bmType As Long
        bmWidth As Long
        bmHeight As Long
        bmWidthBytes As Long
        bmPlanes As Integer
        bmBitsPixel As Integer
        bmBits As Long
End Type
Type BITMAPFILEHEADER
        bfType As Integer
        bfSize As Long
        bfReserved1 As Integer
        bfReserved2 As Integer
        bfOffBits As Long
End Type
Type BITMAPINFOHEADER
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type
Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors As String * 1024 ' Array length is arbitrary; may be changed
End Type

Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Declare Function SetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Declare Function GetWindowWord Lib "user32" (ByVal hwnd As Long, ByVal nIndex As Long) As Integer
Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Declare Function DrawIcon Lib "user32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal hIcon As Long) As Long
Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Declare Function DestroyIcon Lib "user32" (ByVal hIcon As Long) As Long
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const wFlags = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const WM_USER = &H400
Public Const EM_GETLINECOUNT = &HBA
Public Const EM_LINEINDEX = &HBB
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINESCROLL = &HB6
Public Const EM_LINEFROMCHAR = &HC9


Public Const CF_PALETTE = 9
Public Const APICells = 256
Type PALETTEENTRY    '4 Bytes
        peRed As String * 1
        peGreen As String * 1
        peBlue As String * 1
        peFlags As String * 1
End Type
Type LOGPALETTE
  palVersion As Integer 'Windows 3.0 version or higher
  palNumEntries As Integer 'number of color in palette
  palPalEntry(APICells) As PALETTEENTRY 'array of element colors
End Type
Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Public CustPal As LOGPALETTE

Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Declare Function SelectPalette Lib "gdi32" (ByVal hdc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Declare Function RealizePalette Lib "gdi32" (ByVal hdc As Long) As Long
Declare Function OpenClipboard Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Declare Function CloseClipboard Lib "user32" () As Long



Public Sub BaseColorStart()
   Dim I As Long
   
   BaseColcnt = 15: MakeFluent = 1
'   NUseColor(0) = RGB(0, 0, 254)
'   NUseColor(1) = RGB(0, 254, 0)
'   NUseColor(2) = RGB(0, 254, 254)
'   NUseColor(3) = RGB(254, 0, 0)
'   NUseColor(4) = RGB(254, 0, 254)
'   NUseColor(5) = RGB(254, 254, 0)
'   NUseColor(6) = RGB(254, 254, 254)
'   NUseColor(7) = RGB(1, 1, 1)
'   NUseColor(8) = RGB(0, 0, 254)
'   NUseColor(9) = RGB(0, 254, 0)
'   NUseColor(10) = RGB(0, 254, 254)
'   NUseColor(11) = RGB(254, 0, 0)
'   NUseColor(12) = RGB(254, 0, 254)
'   NUseColor(13) = RGB(254, 254, 0)
'   NUseColor(14) = RGB(254, 254, 254)
'   NUseColor(15) = RGB(1, 1, 1)
   NUseColor(0) = 16646144
   NUseColor(1) = 5701206
   NUseColor(2) = 5679358
   NUseColor(3) = 92
   NUseColor(4) = 65278
   NUseColor(5) = 9334784
   NUseColor(6) = 16711422
   NUseColor(7) = 11447982
   NUseColor(8) = 15298186
   NUseColor(9) = 25856
   NUseColor(10) = 16711168
   NUseColor(11) = 254
   NUseColor(12) = 65278
   NUseColor(13) = 11665586
   NUseColor(14) = 16711422

   FloodPalMake BaseColcnt, MakeFluent
   For I = 0 To 15: UseColor(I) = NUseColor(I): Next I ' 15 base colors
   For I = 0 To 255: ColorSet(I) = NColorSet(I): Next I

End Sub

Private Sub CheckRGBByte(ByRef RGBb)
   If RGBb < 0 Then RGBb = 0
   If RGBb > 255 Then RGBb = 255
End Sub

' makes a colorset (240+16) with fluent colors
' fluent from one base color to the next
' "interconnected" or not
Public Sub FloodPalMake(BU As Long, InterConn As Long)
   Dim I As Long, Grp As Long, Grp2 As Long, Stap As Long
   Dim CellsPerGroup As Long, CellNr As Long
   Dim R1 As Long, G1 As Long, B1 As Long
   Dim R2 As Long, G2 As Long, B2 As Long
   Dim Red1 As Long, Green1 As Long, Blue1 As Long
   Dim Temp1 As Long, Temp2 As Long, Temp3 As Long
   Dim Rdis As Single, Gdis As Single, Bdis As Single
   
   CustPal.palVersion = &H300 'Window version
   CustPal.palNumEntries = APICells 'total number of colors
   For I = 0 To 15
      R1 = (QBColor(I) And &HFF)
      G1 = (QBColor(I) And &HFF00&) \ 256
      B1 = (QBColor(I) And &HFF0000) \ 65536
      CustPal.palPalEntry(I).peRed = Chr(R1)
      CustPal.palPalEntry(I).peGreen = Chr(G1)
      CustPal.palPalEntry(I).peBlue = Chr(B1)
      CustPal.palPalEntry(I).peFlags = Chr(0)
      NColorSet(I) = RGB(R1, G1, B1)
   Next I
  
   Select Case BU
   Case 2:
      R1 = (NUseColor(0) And &HFF)
      G1 = (NUseColor(0) And &HFF00&) \ 256
      B1 = (NUseColor(0) And &HFF0000) \ 65536
      R2 = (NUseColor(1) And &HFF)
      G2 = (NUseColor(1) And &HFF00&) \ 256
      B2 = (NUseColor(1) And &HFF0000) \ 65536
      Temp1 = R1: Temp2 = G1: Temp3 = B1
      Rdis = (R1 - R2) / 240: Gdis = (G1 - G2) / 240: Bdis = (B1 - B2) / 240
      For I = 0 To 239
         Red1 = Temp1 - Rdis: Temp1 = Red1
         Green1 = Temp2 - Gdis: Temp2 = Green1
         Blue1 = Temp3 - Bdis: Temp3 = Blue1
         CustPal.palPalEntry(I).peRed = Chr(Red1)
         CustPal.palPalEntry(I).peGreen = Chr(Green1)
         CustPal.palPalEntry(I).peBlue = Chr(Blue1)
         CustPal.palPalEntry(I).peFlags = Chr(0)
         NColorSet(16 + I) = RGB(Red1, Green1, Blue1)
      Next I
   Case Else
      If InterConn = 0 Then Stap = 2 Else Stap = 1
      CellsPerGroup = 240 / BU
      CellNr = 16
      For Grp = 0 To BU - 1 Step Stap
         R1 = (NUseColor(Grp) And &HFF)
         G1 = (NUseColor(Grp) And &HFF00&) \ 256
         B1 = (NUseColor(Grp) And &HFF0000) \ 65536
         If Grp = BU - 1 Then Grp2 = 0 Else Grp2 = Grp + 1
         R2 = (NUseColor(Grp2) And &HFF)
         G2 = (NUseColor(Grp2) And &HFF00&) \ 256
         B2 = (NUseColor(Grp2) And &HFF0000) \ 65536
         Temp1 = R1: Temp2 = G1: Temp3 = B1
         Rdis = (R1 - R2) / (CellsPerGroup * Stap)
         Gdis = (G1 - G2) / (CellsPerGroup * Stap)
         Bdis = (B1 - B2) / (CellsPerGroup * Stap)
         For I = 0 To (CellsPerGroup * Stap) - 1
            Red1 = Temp1 - Rdis: CheckRGBByte Red1: Temp1 = Red1
            Green1 = Temp2 - Gdis: CheckRGBByte Green1: Temp2 = Green1
            Blue1 = Temp3 - Bdis: CheckRGBByte Blue1: Temp3 = Blue1
            CustPal.palPalEntry(CellNr).peRed = Chr(Red1)
            CustPal.palPalEntry(CellNr).peGreen = Chr(Green1)
            CustPal.palPalEntry(CellNr).peBlue = Chr(Blue1)
            CustPal.palPalEntry(CellNr).peFlags = Chr(0)
            NColorSet(CellNr) = RGB(Red1, Green1, Blue1)
            CellNr = CellNr + 1
            If CellNr = 256 Then Exit For
         Next I
         If CellNr = 256 Then Exit For
      Next Grp
   End Select
End Sub

Public Sub APIcls(pic As Control, Kl As Long)
   APIrect pic.hdc, 0, Kl, Kl, 0, 0, pic.ScaleWidth, pic.ScaleHeight
End Sub

'not used
Public Sub APIellipse(hdc As Long, _
                     BorderW As Long, _
                     BorderKl As Long, ByVal FillKl As Long, _
                     iX1 As Long, iY1 As Long, _
                     iX2 As Long, iY2 As Long)
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long
   
   hPn = CreatePen(0, BorderW, BorderKl Xor &H1000000)
   hPnOld = SelectObject(hdc, hPn)
   hBr = CreateSolidBrush(FillKl Xor &H1000000)
   hBrOld = SelectObject(hdc, hBr)
   
   Ellipse hdc, iX1, iY1, iX2, iY2
   
   SelectObject hdc, hBrOld
   DeleteObject hBr
   SelectObject hdc, hPnOld
   DeleteObject hPn
End Sub

Public Sub APIline(hdc As Long, _
                  PenType As Long, _
                  BorderW As Long, BorderKl As Long, _
                  iX1 As Long, iY1 As Long, _
                  iX2 As Long, iY2 As Long)
   Dim Pt As POINTAPI
   Dim hPn As Long, hPnOld As Long
   
   hPn = CreatePen(PenType, BorderW, BorderKl Xor &H1000000)
   hPnOld = SelectObject(hdc, hPn)

   MoveToEx hdc, iX1, iY1, Pt
   LineTo hdc, iX2, iY2

   SelectObject hdc, hPnOld
   DeleteObject hPn
End Sub

Public Sub APIrect(hdc As Long, _
                  BorderW As Long, _
                  BorderKl As Long, FillKl As Long, _
                  iX1 As Long, iY1 As Long, _
                  iX2 As Long, iY2 As Long)
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long

   hPn = CreatePen(0, BorderW, BorderKl Xor &H1000000)
   hPnOld = SelectObject(hdc, hPn)
   hBr = CreateSolidBrush(FillKl Xor &H1000000)
   hBrOld = SelectObject(hdc, hBr)
   
   Rectangle hdc, iX1, iY1, iX2, iY2
   
   SelectObject hdc, hBrOld
   DeleteObject hBr
   SelectObject hdc, hPnOld
   DeleteObject hPn
End Sub

Public Sub APIrrect(hdc As Long, _
                    iPenWidth As Long, _
                    BorderKl As Long, FillKl As Long, _
                    iX1 As Long, iY1 As Long, _
                    iX2 As Long, iY2 As Long, _
                    iRounding As Long)
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long

   If FillKl > -1 Then
      hBr = CreateSolidBrush(FillKl Xor &H1000000)
      hBrOld = SelectObject(hdc, hBr)
      End If
   If BorderKl > -1 Then
      hPn = CreatePen(0, iPenWidth, BorderKl Xor &H1000000)
      hPnOld = SelectObject(hdc, hPn)
      End If
      
   RoundRect hdc, iX1, iY1, iX2, iY2, (iX2 - iX1) * iRounding / 100, (iY2 - iY1) * iRounding / 100
   
   If BorderKl > -1 Then
      SelectObject hdc, hPnOld
      DeleteObject hPn
      End If
   If FillKl > -1 Then
      SelectObject hdc, hBrOld
      DeleteObject hBr
      End If
End Sub

Public Sub APIText(pic As Control, _
                     mX As Long, mY As Long, _
                     txt As String, _
                     Align As Long, Hk As Long, _
                     Kleur As Long)
   Dim H As Long, Wt As Long, I As Long, E As Long
   Dim W, O, U, S, C, OP, CP, Q, PAF
   Dim F As String
   Dim hFnt As Long, hFntOld As Long
   Dim tdX As Long, tdY As Long
   Dim mX1 As Long, mY1 As Long
   
   SetTextColor pic.hdc, Kleur&
   If InStr(txt, vbCrLf) Then
      RC.Left = 0: RC.Top = 0
      DrawText pic.hdc, txt, Len(txt), RC, Align Or DT_CALCRECT
      OffsetRect RC, mX - RC.Right \ 2, mY - RC.Bottom \ 2
      DrawText pic.hdc, txt, Len(txt), RC, Align
      Else
      Dim TM As TEXTMETRIC
      Dim Sz As Size
      GetTextMetrics pic.hdc, TM
      H = TM.tmHeight
      Wt = TM.tmWeight
      I = TM.tmItalic 'Asc(TM.tmItalic)
      F$ = String(128, " ")
      GetTextFace pic.hdc, 128, F$
      E = Hk
      hFnt = CreateFont(H, W, E, O, Wt, I, U, S, C, OP, CP, Q, PAF, F$)
      hFntOld = SelectObject(pic.hdc, hFnt)
      GetTextExtentPoint pic.hdc, txt, Len(txt), Sz
      Select Case Hk
      Case 0
         tdX = Sz.cX: tdY = Sz.cY
         mX1 = mX - tdX \ 2: mY1 = mY - tdY \ 2
      Case 900
         tdX = Sz.cY: tdY = Sz.cX
         mX1 = mX - tdX \ 2: mY1 = mY + tdY \ 2: tdY = -tdY
      Case -900
         tdX = Sz.cY: tdY = Sz.cX
         mX1 = mX + tdX \ 2: tdX = -tdX: mY1 = mY - tdY \ 2
      Case 1800
         tdX = Sz.cX: tdY = Sz.cY
         mX1 = mX + tdX \ 2: tdX = -tdX: mY1 = mY + tdY \ 2: tdY = -tdY
      End Select
      TextOut pic.hdc, mX1, mY1, txt, Len(txt)
      SelectObject pic.hdc, hFntOld
      DeleteObject hFnt
      End If
End Sub

Public Sub AutoScroll(kadI As Control, _
                      ByVal X As Long, ByVal Y As Long, _
                      HS As Control, VS As Control)
   Dim Unit As Long
   Dim nHS As Long, nVS As Long
   
   Unit = 15 * PixSze * 1 ' pixels to twips enlarged by zoomfactor
   X = X * Unit: Y = Y * Unit
   If HS.Visible = False Then GoTo ASverticaal:
   Unit = Unit * 2
   If X - HS.Value >= kadI.Width Then
      nHS = HS.Value + Unit
      If nHS > HS.Max Then nHS = HS.Max
      HS.Value = nHS
      End If
   If X <= HS.Value Then
      nHS = HS.Value - Unit
      If nHS < HS.Min Then nHS = HS.Min
      HS.Value = nHS
      End If
ASverticaal:
   If VS.Visible = False Then Exit Sub
   If Y - VS.Value >= kadI.Height Then
      nVS = VS.Value + Unit
      If nVS > VS.Max Then nVS = VS.Max
      VS.Value = nVS
      End If
   If Y <= VS.Value Then
      nVS = VS.Value - Unit
      If nVS < VS.Min Then nVS = VS.Min
      VS.Value = nVS
      End If
End Sub

' the kad (frame) pictures must be set to twips mode for this
' routine to work
Public Sub CheckScrolls(pic As Control, _
                        kadI As Control, kadO As Control, _
                        HS As Control, VS As Control)
   kadI.Width = kadO.Width - 60
   VS.Left = kadI.Width - VS.Width + 15
   HS.Width = kadI.Width - VS.Width
   
   kadI.Height = kadO.Height - 60
   HS.Top = kadI.Height - HS.Height + 15
   VS.Height = kadI.Height - HS.Height
   
   If pic.Width > kadI.Width - 90 Then
      kadI.Height = kadO.Height - HS.Height - 90: VS.Height = kadI.Height
      HS.Max = pic.Width - kadI.Width + 45: HS.Visible = True
      Else
      VS.Height = kadI.Height
      HS.Value = HS.Min: HS.Visible = False
      End If
   If pic.Height > kadI.Height - 90 Then
      kadI.Width = kadO.Width - VS.Width - 90: HS.Width = kadI.Width
      VS.Max = pic.Height - kadI.Height + 45: VS.Visible = True
      Else
      HS.Width = kadI.Width
      VS.Value = VS.Min: VS.Visible = False
      End If
   If pic.Width > kadI.Width - 90 And pic.Height > kadI.Height - 90 Then
      kadI.Height = kadO.Height - HS.Height - 90: VS.Height = kadI.Height
      HS.Max = pic.Width - kadI.Width + 45: HS.Visible = True
      kadI.Width = kadO.Width - VS.Width - 90: HS.Width = kadI.Width
      VS.Max = pic.Height - kadI.Height + 45: VS.Visible = True
      End If
End Sub

Function FixPath(ByVal P As Variant) As String
   If Right(P, 1) = "\" Then FixPath = P Else FixPath = P & "\"
End Function

Public Sub Flood8b(pic As Control, _
                   StartColor As Long, EndColor As Long, _
                   ByVal Stijl As Long)
   Dim Pt As POINTAPI
   Dim KlDis As Long          ' EndColor(Index)-StartColor(Index)
   Dim ID As Long             ' color ID counter
   Dim I As Long              ' counter
   Dim St As Single           ' after how much of I (pixels) a color ID has to change
   Dim StK As Single          ' with which amount colorID's are changed
   Dim W As Long, H As Long   ' Width-Height
   Dim D As Long              ' distance in pixels
   Dim X As Long, Y As Long
   Dim XX1 As Long, YY1 As Long
   Dim XX2 As Long, YY2 As Long
   
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long
   
   If StartColor < 16 Then StartColor = 16
   If EndColor < 16 Then EndColor = 255
   KlDis = EndColor - StartColor
   ID = StartColor - 16
   APIcls pic, 7
   '
   Select Case Stijl
   
   Case FT_LEFTRIGHT, FT_LEFTRIGHT2  ' West-East
   X = 0: Y = -1
   W = pic.ScaleWidth: H = pic.ScaleHeight + 1
   If Stijl = FT_LEFTRIGHT2 Then D = W / 2 Else D = W
   Select Case D
     Case Is > Abs(KlDis): St = (Abs(D / KlDis)): StK = 1
     Case Is < Abs(KlDis): St = 1: StK = (Abs(KlDis / D))
     Case Else: St = 1: StK = 1
   End Select
   If St < 1 Then St = 1
   For I = 0 To W - 1
     hPn = CreatePen(0, 0, ColorSet(16 + ID) Xor &H2000000) ' Solid(0), Standard width(0) so 1 pixel
     hPnOld = SelectObject(pic.hdc, hPn)  ' select pen
     MoveToEx pic.hdc, X + I, Y, Pt       '
     LineTo pic.hdc, X + I, Y + H         ' use pen
     SelectObject pic.hdc, hPnOld         ' (re)set to prev. pen
     DeleteObject hPn                     ' remove new used pen
     If I >= W \ 2 - 1 And Stijl = FT_LEFTRIGHT2 Then
        KlDis = -KlDis: Stijl = FT_LEFTRIGHT
        If W \ 2 <> W / 2 Then ID = (240 + ID + StK * Sgn(KlDis)) Mod 240
        Else
        If I Mod St = 0 Then ID = (240 + ID + StK * Sgn(KlDis)) Mod 240
        End If
   Next I
   
   Case FT_UPDOWN, FT_UPDOWN2  'Nord-South
   X = 0: Y = 0
   W = pic.ScaleWidth: H = pic.ScaleHeight
   If Stijl = FT_UPDOWN2 Then D = H / 2 Else D = H
   Select Case D
     Case Is > Abs(KlDis): St = (Abs(D / KlDis)): StK = 1
     Case Is < Abs(KlDis): St = 1: StK = (Abs(KlDis / D))
     Case Else: St = 1: StK = 1
   End Select
   If St < 1 Then St = 1
   For I = 0 To H - 1
     hPn = CreatePen(0, 0, ColorSet(16 + ID) Xor &H2000000)
     hPnOld = SelectObject(pic.hdc, hPn)
     MoveToEx pic.hdc, X, Y + I, Pt
     LineTo pic.hdc, X + W, Y + I
     SelectObject pic.hdc, hPnOld
     DeleteObject hPn
     If I >= H \ 2 And Stijl = FT_UPDOWN2 Then
        KlDis = -KlDis: Stijl = FT_UPDOWN
        If H \ 2 <> H / 2 Then ID = (240 + ID + StK * Sgn(KlDis)) Mod 240
        Else
        If I Mod St = 0 Then ID = (240 + ID + StK * Sgn(KlDis)) Mod 240
        End If
   Next I
   
   Case FT_ULDR, FT_ULDR2  ' diagonal
   X = 0: Y = 0
   W = pic.ScaleWidth: H = pic.ScaleHeight
   If Stijl = FT_ULDR2 Then D = (W + H) \ 2 Else D = W + H
   Select Case D
     Case Is > Abs(KlDis): St = (Abs(D / KlDis)): StK = 1
     Case Is < Abs(KlDis): St = 1: StK = (Abs(KlDis / D))
     Case Else: St = 1: StK = 1
   End Select
   If St < 1 Then St = 1
   I = 0
   Do While I < W + H
     
     If I < W Then XX2 = X + I: YY2 = Y Else YY2 = YY2 + 1
     If I < H Then YY1 = Y + I: XX1 = X Else XX1 = XX1 + 1
     hPn = CreatePen(0, 2, ColorSet(16 + ID) Xor &H2000000)
     hPnOld = SelectObject(pic.hdc, hPn)
     MoveToEx pic.hdc, XX1, YY1, Pt
     LineTo pic.hdc, XX2, YY2
     SelectObject pic.hdc, hPnOld
     DeleteObject hPn
     I = I + 1
     If I = D And Stijl = FT_ULDR2 Then
        KlDis = -KlDis: Stijl = FT_ULDR
        Else
        If I Mod St = 0 Then ID = (240 + ID + StK * Sgn(KlDis)) Mod 240
        End If
   Loop
   
   Case FT_DLUR, FT_DLUR2  'diagonal2
   X = 0: Y = -1
   W = pic.ScaleWidth: H = pic.ScaleHeight + 1
   If Stijl = FT_DLUR2 Then D = (W + H) \ 2 Else D = W + H
   Select Case D
     Case Is > Abs(KlDis): St = (Abs(D / KlDis)): StK = 1
     Case Is < Abs(KlDis): St = 1: StK = (Abs(KlDis / D))
     Case Else: St = 1: StK = 1
   End Select
   If St < 1 Then St = 1
   I = 0
   Do While I < W + H
     If I < W + 1 Then XX2 = X + (W - I): YY2 = Y Else YY2 = YY2 + 1
     If I < H Then YY1 = Y + I: XX1 = X + W Else XX1 = XX1 - 1
     hPn = CreatePen(0, 0, ColorSet(16 + ID) Xor &H2000000)
     hPnOld = SelectObject(pic.hdc, hPn)
     MoveToEx pic.hdc, XX1 - 1, YY1, Pt
     LineTo pic.hdc, XX2 - 1, YY2
     SelectObject pic.hdc, hPnOld
     DeleteObject hPn
     I = I + 1
     If I = D And Stijl = FT_DLUR2 Then
        KlDis = -KlDis: Stijl = FT_DLUR
        Else
        If I Mod St = 0 Then ID = (240 + ID + StK * Sgn(KlDis)) Mod 240
        End If
   Loop
   
   Case FT_CIRCLE  ' circles
   W = pic.ScaleWidth: H = pic.ScaleHeight
   X = W / 2: Y = H / 2
   D = Sqr(X * X + Y * Y)
   Select Case D
     Case Is > Abs(KlDis): St = (Abs(D / KlDis)): StK = 1
     Case Is < Abs(KlDis): St = 1: StK = (Abs(KlDis / D))
     Case Else: St = 1: StK = 1
   End Select
   If St < 1 Then St = 1
   XX1 = X + D * Cos(PI)
   YY1 = Y + D * Sin(PI * 3 / 2)
   XX2 = X + D * Cos(0)
   YY2 = Y + D * Sin(PI * 1 / 2)
   While XX1 < XX2 And YY1 < YY2
     hPn = CreatePen(0, 2, ColorSet(16 + ID) Xor &H2000000)
     hPnOld = SelectObject(pic.hdc, hPn)
     Ellipse pic.hdc, XX1, YY1, XX2, YY2
     SelectObject pic.hdc, hPnOld
     DeleteObject hPn
     XX1 = XX1 + 1: YY1 = YY1 + 1
     XX2 = XX2 - 1: YY2 = YY2 - 1
     I = I + 1
     If I Mod St = 0 Then ID = (240 + ID + StK * Sgn(KlDis)) Mod 240
   Wend
   
   Case FT_SQUARE ' rectangle
   X = 0: Y = 0
   W = pic.ScaleWidth: H = pic.ScaleHeight
   
   If H > D Then D = H / 2 Else D = W / 2
   Select Case D
     Case Is > Abs(KlDis): St = (Abs(D / KlDis)): StK = 1
     Case Is < Abs(KlDis): St = 1: StK = (Abs(KlDis / D))
     Case Else: St = 1: StK = 1
   End Select
   If St < 1 Then St = 1
   I = 0
   XX1 = 0: YY1 = 0: XX2 = W: YY2 = H
   While XX1 < XX2 And YY1 < YY2
     hPn = CreatePen(0, 2, ColorSet(16 + ID) Xor &H2000000)
     hPnOld = SelectObject(pic.hdc, hPn)
     Rectangle pic.hdc, XX1, YY1, XX2, YY2
     SelectObject pic.hdc, hPnOld
     DeleteObject hPn
     XX1 = XX1 + 1: YY1 = YY1 + 1
     XX2 = XX2 - 1: YY2 = YY2 - 1
     I = I + 1
     If I Mod St = 0 Then ID = (240 + ID + StK * Sgn(KlDis)) Mod 240
   Wend
   
   End Select
   If pic.AutoRedraw = True Then pic.Refresh

End Sub

' obj must be one who knows the .Line method
Public Sub BevelObject(obj As Object, _
                     X1 As Long, Y1 As Long, _
                     X2 As Long, Y2 As Long, _
                     ByRef Bevel As Long)
   Dim ColorUpper As Long
   Dim ColorLower As Long
   
   If Bevel = 0 Then
      ColorUpper = &H80000014
      ColorLower = &H80000010
      Else
      ColorUpper = &H80000010
      ColorLower = &H80000014
      End If
   obj.Line (X1, Y1)-(X2, Y1), ColorUpper
   obj.Line (X1, Y1)-(X1, Y2), ColorUpper
   obj.Line (X2, Y1)-(X2, Y2), ColorLower
   obj.Line (X1, Y2)-(X2, Y2), ColorLower
End Sub

Public Function IsColorID(Color As Long, StartID As Long) As Integer
   Dim I As Long
   
   For I = StartID To 255
   If ColorSet(I) = Color Then Exit For
   Next I
   IsColorID = I
End Function

Public Sub Pause(ByVal Pze As Long)
   Dim mTime As Variant
   mTime = Timer
   While Timer - mTime < Pze / 1000: DoEvents: Wend
End Sub

Public Sub DrawFormula(pic As Control, _
                      X1, Y1, dX, dY, _
                      Pts, GPlus, _
                      Angle, Kl&, Vul)
   Dim hdc As Long
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long
   Dim xs As Long, ys As Long
   Dim xc As Long, yc As Long
   ReDim Pt(Pts) As POINTAPI
   Dim P As Long, G As Long
   
   hdc = pic.hdc
   
   On Error GoTo TFrmFout:
   dX = dX - 1: dY = dY - 1
   xs = dX / 2: xc = X1 + xs
   ys = dY / 2: yc = Y1 + ys
   
   For P = 0 To Pts - 1
      G = (G + GPlus) Mod 360
      Pt(P).X = xc + xs * Cos((G + Angle) * PI / 180)
      Pt(P).Y = yc + ys * Sin((G + Angle) * PI / 180)
   Next P
   hPn = CreatePen(0, PenWidth, Kl& Xor &H2000000)
   hPnOld = SelectObject(hdc, hPn)
   hBr = CreateSolidBrush(Kl& Xor &H20000000)
   hBrOld = SelectObject(hdc, hBr)
   If Vul = 1 Then
      Polygon hdc, Pt(0), Pts
      Else
      Polyline hdc, Pt(0), Pts
      End If
   SelectObject hdc, hBrOld
   DeleteObject hBr
   SelectObject hdc, hPnOld
   DeleteObject hPn
   
   If pic.AutoRedraw = True Then pic.Refresh
TFrmEinde:
   Exit Sub
TFrmFout:
   MsgBox "Fout:" & Str(Err) & vbCrLf & Error$
   Resume TFrmEinde:
End Sub

Public Sub DrawArrow(ByVal hdc As Long, ByVal cX As Long, ByVal cY As Long, _
                        ByVal Angle As Long, ByVal dX As Long, ByVal dY As Long, ByVal ArrWidth As Long, _
                        ByVal HeadW As Single, ByVal DoubleArr As Boolean)

   Dim Xc1 As Long, Yc1 As Long
   Dim Xc2 As Long, Yc2 As Long
   Dim Xc3 As Long, Yc3 As Long
   Dim AngleXA As Single, AngleYA As Single  ' back-side
   Dim AngleXL As Single, AngleYL As Single  ' turn to the left
   Dim AngleXR As Single, AngleYR As Single  ' turn to the right
   Dim AngleXV As Single, AngleYV As Single  ' fore-side
   Dim StraalAx As Long, StraalAy As Long
   Dim Pt(12) As POINTAPI, AaPt As Long
   Dim Deg As Single
   
   Deg = (Atn(1) * 4) / 180
   Angle = (360 + 180 - Angle) Mod 360
   AngleXA = Cos(Angle * Deg):         AngleYA = Sin(Angle * Deg)
   AngleXV = Cos((Angle + 180) * Deg): AngleYV = Sin((Angle + 180) * Deg)
   AngleXL = Cos((Angle - 90) * Deg):  AngleYL = Sin((Angle - 90) * Deg)
   AngleXR = Cos((Angle + 90) * Deg):  AngleYR = Sin((Angle + 90) * Deg)
   StraalAx = dX: StraalAy = dY
   If ArrWidth > dX Then ArrWidth = dX
   If ArrWidth > dY Then ArrWidth = dY
   If HeadW < ArrWidth Then HeadW = ArrWidth
   If ArrWidth > HeadW Then ArrWidth = HeadW
   If dX < HeadW Then
      HeadW = dX
      dX = 0
      Else
      dX = dX - HeadW
      End If
   If dY < HeadW Then
      HeadW = dY
      dY = 0
      Else
      dY = dY - HeadW
      End If
      
   If DoubleArr = True Then StraalAx = dX: StraalAy = dY
   
   Xc1 = cX + dX * AngleXV
   Yc1 = cY + dY * AngleYV
   Xc2 = cX + StraalAx * AngleXA
   Yc2 = cY + StraalAy * AngleYA
   Xc3 = cX - dX * AngleXV
   Yc3 = cY - dY * AngleYV
   
   Pt(0).X = Xc2 + ArrWidth * AngleXL
   Pt(0).Y = Yc2 + ArrWidth * AngleYL
   Pt(1).X = Xc1 + ArrWidth * AngleXL
   Pt(1).Y = Yc1 + ArrWidth * AngleYL
   Pt(2).X = Xc1 + HeadW * AngleXL
   Pt(2).Y = Yc1 + HeadW * AngleYL
   Pt(3).X = Xc1 + HeadW * AngleXV
   Pt(3).Y = Yc1 + HeadW * AngleYV
   Pt(4).X = Xc1 + HeadW * AngleXR
   Pt(4).Y = Yc1 + HeadW * AngleYR
   Pt(5).X = Xc1 + ArrWidth * AngleXR
   Pt(5).Y = Yc1 + ArrWidth * AngleYR
   Pt(6).X = Xc2 + ArrWidth * AngleXR
   Pt(6).Y = Yc2 + ArrWidth * AngleYR
   If DoubleArr = False Then
      AaPt = 8
      Pt(7).X = Xc2 + ArrWidth * AngleXL
      Pt(7).Y = Yc2 + ArrWidth * AngleYL
      Else
      AaPt = 12
      Pt(7).X = Xc3 - ArrWidth * AngleXL
      Pt(7).Y = Yc3 - ArrWidth * AngleYL
      Pt(8).X = Xc3 - HeadW * AngleXL
      Pt(8).Y = Yc3 - HeadW * AngleYL
      Pt(9).X = Xc3 - HeadW * AngleXV
      Pt(9).Y = Yc3 - HeadW * AngleYV
      Pt(10).X = Xc3 - HeadW * AngleXR
      Pt(10).Y = Yc3 - HeadW * AngleYR
      Pt(11).X = Xc3 - ArrWidth * AngleXR
      Pt(11).Y = Yc3 - ArrWidth * AngleYR
      End If
   Polygon hdc, Pt(0), AaPt
End Sub



