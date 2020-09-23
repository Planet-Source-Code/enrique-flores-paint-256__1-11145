VERSION 5.00
Begin VB.Form frmPal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Make/Set 240 Index Color Pallet"
   ClientHeight    =   3405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5715
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   227
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   381
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      DownPicture     =   "frmPal.frx":0000
      Height          =   540
      Left            =   1830
      Picture         =   "frmPal.frx":03E6
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2700
      Width           =   1785
   End
   Begin VB.CommandButton cmdCancel 
      DownPicture     =   "frmPal.frx":07CC
      Height          =   540
      Left            =   3720
      Picture         =   "frmPal.frx":0C6E
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   2700
      Width           =   1815
   End
   Begin VB.ComboBox cmbBaseColcnt 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   2550
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   150
      Width           =   675
   End
   Begin VB.CheckBox chkFluent 
      Caption         =   "Fluent"
      ForeColor       =   &H00008000&
      Height          =   315
      Left            =   4440
      TabIndex        =   9
      Top             =   165
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.PictureBox picPal 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      DragIcon        =   "frmPal.frx":1110
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1845
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   240
      TabIndex        =   8
      Top             =   915
      Width           =   3630
   End
   Begin VB.PictureBox picBcol 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      DragIcon        =   "frmPal.frx":141A
      DrawMode        =   6  'Mask Pen Not
      DrawWidth       =   2
      ForeColor       =   &H80000008&
      Height          =   165
      Index           =   0
      Left            =   1785
      ScaleHeight     =   9
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   9
      TabIndex        =   7
      ToolTipText     =   "Basiskleur"
      Top             =   615
      Width           =   165
   End
   Begin VB.PictureBox picSel 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000014&
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   4485
      ScaleHeight     =   630
      ScaleWidth      =   960
      TabIndex        =   6
      Top             =   1725
      Width           =   990
   End
   Begin VB.CommandButton cmdSave 
      Appearance      =   0  'Flat
      Caption         =   "&Save"
      Height          =   420
      Left            =   60
      MaskColor       =   &H8000000F&
      TabIndex        =   5
      ToolTipText     =   "Palet bewaren"
      Top             =   585
      Width           =   1530
   End
   Begin VB.CommandButton cmdOpen 
      Appearance      =   0  'Flat
      Caption         =   "&Open"
      Height          =   420
      Left            =   60
      MaskColor       =   &H8000000F&
      TabIndex        =   4
      ToolTipText     =   "Palet openen"
      Top             =   105
      Width           =   1530
   End
   Begin VB.CommandButton cmdMakePal 
      Appearance      =   0  'Flat
      Caption         =   "&Make Pallet"
      Default         =   -1  'True
      Height          =   870
      Left            =   45
      MaskColor       =   &H8000000F&
      TabIndex        =   3
      ToolTipText     =   "Palet aanmaken"
      Top             =   1635
      Width           =   1545
   End
   Begin VB.HScrollBar hsRGB 
      Height          =   210
      Index           =   0
      LargeChange     =   8
      Left            =   1830
      Max             =   255
      TabIndex        =   2
      Top             =   1725
      Width           =   2295
   End
   Begin VB.HScrollBar hsRGB 
      Height          =   210
      Index           =   1
      LargeChange     =   8
      Left            =   1830
      Max             =   255
      TabIndex        =   1
      Top             =   1950
      Width           =   2295
   End
   Begin VB.HScrollBar hsRGB 
      Height          =   210
      Index           =   2
      LargeChange     =   8
      Left            =   1830
      Max             =   255
      TabIndex        =   0
      Top             =   2175
      Width           =   2295
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Use"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   0
      Left            =   1860
      TabIndex        =   15
      Top             =   195
      Width           =   495
   End
   Begin VB.Label lbl 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "base colors"
      ForeColor       =   &H00008000&
      Height          =   255
      Index           =   1
      Left            =   3255
      TabIndex        =   14
      Top             =   195
      Width           =   1155
   End
   Begin VB.Label lblRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FF0000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   225
      Index           =   2
      Left            =   4110
      TabIndex        =   13
      Top             =   2160
      Width           =   390
   End
   Begin VB.Label lblRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H0000FF00&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Index           =   1
      Left            =   4110
      TabIndex        =   12
      Top             =   1935
      Width           =   390
   End
   Begin VB.Label lblRGB 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Index           =   0
      Left            =   4110
      TabIndex        =   11
      Top             =   1725
      Width           =   390
   End
   Begin VB.Line linBCol 
      Index           =   0
      X1              =   123
      X2              =   123
      Y1              =   57
      Y2              =   63
   End
   Begin VB.Shape shpSelect 
      BorderColor     =   &H8000000E&
      Height          =   195
      Left            =   1770
      Top             =   600
      Width           =   195
   End
End
Attribute VB_Name = "frmPal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Loading As Boolean
Dim BColNr As Integer ' base color nr.
Dim DragColor As Long
Dim DragIndex As Long
Dim Offset As Long

Private Sub ShowRGB(ByVal Color As Long)
   Dim lngRed As Long, lngGreen As Long, lngBlue As Long
   Dim I As Long
   
   lngRed = (Color And &HFF)
   lngGreen = (Color And &HFF00&) \ 256
   lngBlue = (Color And &HFF0000) \ 65536
   lblRGB(0).Caption = lngRed
   lblRGB(1).Caption = lngGreen
   lblRGB(2).Caption = lngBlue
   hsRGB(0).Value = lngRed
   hsRGB(1).Value = lngGreen
   hsRGB(2).Value = lngBlue
   picSel.BackColor = RGB(lngRed, lngGreen, lngBlue)
   picBcol(BColNr).BackColor = RGB(lngRed, lngGreen, lngBlue)
End Sub
' position picBcol's and set their colors
Private Sub ShowBaseColors()
   Dim I As Long, X As Long, Dis As Long
   Dim LinesLeft As Long, LinesTop As Long
   
   For I = 0 To 15
      linBCol(I).Visible = IIf(I < NBaseColcnt, True, False)
      picBcol(I).Visible = linBCol(I).Visible
   Next I
   
   Dis = 240 / NBaseColcnt
   LinesLeft = picPal.Left
   LinesTop = picPal.Top
   For I = 0 To NBaseColcnt - 1
      X = LinesLeft + I * Dis + 1
      linBCol(I).X1 = X: linBCol(I).X2 = X
      linBCol(I).Y1 = LinesTop - 6: linBCol(I).Y2 = LinesTop
      picBcol(I).Left = X - 5: picBcol(I).Top = LinesTop - 16
   Next I
   If chkFluent.Value = 0 Then
      LinesTop = picPal.Top + 20
      For I = 1 To NBaseColcnt - 1 Step 2
         X = LinesLeft + (I + 1) * Dis
         linBCol(I).X1 = X: linBCol(I).X2 = X
         linBCol(I).Y1 = LinesTop: linBCol(I).Y2 = LinesTop + 6
         picBcol(I).Left = X - 5: picBcol(I).Top = LinesTop + 6
      Next I
      End If
   picBcol_click BColNr
   cmdMakePal.Enabled = True
End Sub

' X = color index on X-axe
Private Function IsBaseColor(ByVal X As Long) As Boolean
   Dim I As Long
   
   For I = 0 To BaseColcnt - 1
     If linBCol(I).Visible = False Then Exit For ' no more base colors
     If X = linBCol(I).X1 Then
        IsBaseColor = True
        BColNr = I + 1 'return value is found base color
        Exit For
        End If
   Next I
End Function

Private Sub PalInPicPal()
   Dim Ret As Long
   Dim hPal As Long
   
   Clipboard.Clear
   Ret = OpenClipboard(Me.hwnd)
   If Ret = 0 Then MsgBox "clipboard error": End
   hPal = CreatePalette(CustPal)
   Ret = SetClipboardData(CF_PALETTE, hPal)
   Ret = CloseClipboard()
   picPal.Picture = Clipboard.GetData(CF_PALETTE)
   Ret = DeleteObject(hPal)
End Sub

' extract a colorset from a bitmap file
' returns the number of extracted colors
Function PALfromBMP(ByVal file As String)
   Dim Ch As Long, I As Long, k As Long
   Dim BFh As BITMAPFILEHEADER
   Dim BIh As BITMAPINFO ' 40 bytes+colors
   Dim PalBytes As Long, ColCount As Long
   Dim Pal As String
   Dim R As Long, G As Long, b As Long

   Ch = FreeFile
   Open file For Binary As Ch
   Get #Ch, 1, BFh
   Get #Ch, , BIh
   PalBytes = BFh.bfOffBits - BIh.bmiHeader.biSize - 14
   If PalBytes > 0 Then
      Pal = Left(BIh.bmiColors, PalBytes)
      If BIh.bmiHeader.biClrUsed = 0 Then
         ColCount = 2 ^ BIh.bmiHeader.biBitCount
         Else
         ColCount = BIh.bmiHeader.biClrUsed
         End If
      If ColCount <= 256 Then
         For I = 1 To Len(Pal) - 1 Step 4
            b = Asc(Mid(Pal, I, 1))
            G = Asc(Mid(Pal, I + 1, 1))
            R = Asc(Mid(Pal, I + 2, 1))
            NColorSet(k) = RGB(R, G, b)
            CustPal.palPalEntry(k).peRed = Chr(R)
            CustPal.palPalEntry(k).peGreen = Chr(G)
            CustPal.palPalEntry(k).peBlue = Chr(b)
            k = k + 1: If k >= APICells Then Exit For
         Next I
         End If
      End If
   Close Ch
   PALfromBMP = k

End Function

Private Sub ShowPallet()
   Dim I As Long
   picPal.Cls
   For I = 16 To 255
      picPal.Line (I - 16, 0)-(I - 16, 18), NColorSet(I)
   Next I
End Sub

Private Sub chkFluent_Click()
   ShowBaseColors
End Sub

Private Sub cmbBaseColcnt_Click()
   Dim I As Long
   
   picBcol_click 0
   NBaseColcnt = Val(cmbBaseColcnt.Text)
   For I = 0 To 15
      picBcol(I).Visible = IIf(I > NBaseColcnt, False, True)
   Next I
   ShowBaseColors
   ShowPallet
End Sub

Private Sub cmdMakePal_click()
   Dim I As Long
   
   Me.MousePointer = vbHourglass
   
   For I = 0 To 15: NUseColor(I) = picBcol(I).BackColor: Next I
   FloodPalMake NBaseColcnt, chkFluent.Value
   PalInPicPal
   ShowPallet
   
   Me.MousePointer = vbDefault
   cmdMakePal.Enabled = False
End Sub

' 4 ways to store colorsets
Private Sub cmdSave_Click()
   Dim Ch As Long, I As Long, Nr As Long
   Dim Col As Long, ColCount As Long
   Dim Filter As String, txt As String
   
   Filter = "Base colors (*.BKL)" & Chr(0) & "*.BKL" & Chr(0)
   Filter = Filter & "JASC - PAL  (*.PAL)" & Chr(0) & "*.PAL" & Chr(0)
   Filter = Filter & "Fractint (*.MAP)" & Chr(0) & "*.MAP" & Chr(0)
   Filter = Filter & "Bitmap (*.BMP)" & Chr(0) & "*.BMP" & Chr(0)
   
   CD_File.hWndOwner = Me.hwnd
   CD_File.Filter = Filter
   CD_File.FilterIndex = 0
   CD_File.FileName = "*.BKL"
   CD_File.DialogTitle = "Save File"
   CD_File.Flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
   CD_File.DefaultExt = ".BKL"
   On Error Resume Next
   CD_File.ShowSave
   If Err <> 0 Then MsgBox Err.Description: Exit Sub

   Select Case UCase(Right(CD_File.FileName, 3))
   
   Case "BKL"
      Ch = FreeFile
      Open CD_File.FileName For Output As Ch
         Write #Ch, cmbBaseColcnt.ListIndex, chkFluent.Value
         For I = 0 To BaseColcnt - 1
            Col = picBcol(I).BackColor
            Write #Ch, Col
         Next I
      Close Ch

   Case "PAL"
      txt = "JASC-PAL" & vbCrLf
      txt = txt & "0100" & vbCrLf
      txt = txt & Format(APICells) & vbCrLf
      For I = Offset To APICells - 1
         txt = txt & Format(Asc(CustPal.palPalEntry(I).peRed))
         txt = txt & Str(Asc(CustPal.palPalEntry(I).peGreen))
         txt = txt & Str(Asc(CustPal.palPalEntry(I).peBlue))
         txt = txt & vbCrLf
      Next I
      If Offset > 0 Then
         For I = 0 To Offset - 1
            txt = txt & Format(Asc(CustPal.palPalEntry(APICells - Offset + I).peRed))
            txt = txt & Str(Asc(CustPal.palPalEntry(APICells - Offset + I).peGreen))
            txt = txt & Str(Asc(CustPal.palPalEntry(APICells - Offset + I).peBlue))
            txt = txt & vbCrLf
         Next I
         End If
      Ch = FreeFile
      Open CD_File.FileName For Output As Ch
      Print #Ch, txt
      Close Ch

   Case "MAP"
      For I = Offset To APICells - 1
         txt = txt & Format(Asc(CustPal.palPalEntry(I).peRed))
         txt = txt & Str(Asc(CustPal.palPalEntry(I).peGreen))
         txt = txt & Str(Asc(CustPal.palPalEntry(I).peBlue))
         txt = txt & vbCrLf
      Next I
      If Offset > 0 Then
         For I = 0 To Offset - 1
            txt = txt & Format(Asc(CustPal.palPalEntry(I).peRed))
            txt = txt & Str(Asc(CustPal.palPalEntry(I).peGreen))
            txt = txt & Str(Asc(CustPal.palPalEntry(I).peBlue))
            txt = txt & vbCrLf
         Next I
         End If
      Ch = FreeFile
      Open CD_File.FileName For Output As Ch
      Print #Ch, txt
      Close Ch

   Case "BMP"
      SavePicture picPal.Image, CD_File.FileName
   Case Else
      MsgBox "Unknown or missing file-extention"
   End Select
   
End Sub

Private Sub cmdCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   OK = False
   Unload Me 'Hide
End Sub

Private Sub cmdOpen_click()
   Dim Ch As Long, I As Long, Nr As Long
   Dim cL As Long, cV As Long, Col As Long, ColCount As Long
   Dim R As Long, G As Long, b As Long
   Dim Filter As String, txt As String
   Dim pos1 As Long, pos2 As Long, pos3 As Long
   
   Filter = "Base colors (*.BKL)" & Chr(0) & "*.BKL" & Chr(0)
   Filter = Filter & "JASC - PAL  (*.PAL)" & Chr(0) & "*.PAL" & Chr(0)
   Filter = Filter & "Fractint (*.MAP)" & Chr(0) & "*.MAP" & Chr(0)
   Filter = Filter & "Bitmap (*.BMP)" & Chr(0) & "*.BMP" & Chr(0)
   
   CD_File.hWndOwner = Me.hwnd
   CD_File.Filter = Filter
   CD_File.FilterIndex = 0
   CD_File.FileName = ""
   CD_File.DialogTitle = "Open File"
   CD_File.Flags = OFN_FILEMUSTEXIST
   On Error Resume Next
   CD_File.ShowOpen
   If Err <> 0 Then MsgBox Err.Description: Exit Sub
   
   Select Case UCase(Right(CD_File.FileName, 3))
   Case "BKL"
      Ch = FreeFile
      Open CD_File.FileName For Input As Ch
        Input #Ch, cL, cV
        cmbBaseColcnt.ListIndex = cL
        chkFluent.Value = cV
        For I = 0 To NBaseColcnt - 1
          Input #Ch, Col
          picBcol(I).BackColor = Col
        Next I
      Close Ch
      If NBaseColcnt < 16 Then
         For I = NBaseColcnt To 15
            picBcol(I).BackColor = QBColor(15)
         Next I
         End If
      For I = 0 To 15: NUseColor(I) = picBcol(I).BackColor: Next I
      FloodPalMake NBaseColcnt, chkFluent.Value
      cmdMakePal.Enabled = False

   Case "PAL"
      Ch = FreeFile
      Open CD_File.FileName For Input As Ch
         Line Input #Ch, txt
         If UCase(Left(txt, 8)) <> "JASC-PAL" Then
            MsgBox "This is not a 'JASC-PAL' pallet file.":  Exit Sub
            End If
         Line Input #Ch, txt ' "0100"
         Line Input #Ch, txt
         ColCount = Val(Left(txt, 3))
         If ColCount = 0 Then MsgBox "No colors present (3th row)?!": Exit Sub
         Nr = 0
         While Not EOF(Ch) And Nr < APICells
            Line Input #Ch, txt
            pos1 = InStr(1, txt, " ")
            pos2 = InStr(pos1 + 1, txt, " ")
            pos3 = InStr(pos2 + 1, txt, " ")
            R = Val(Left(txt, pos1 - 1))
            G = Val(Mid(txt, pos1 + 1, pos2 - pos1))
            If pos3 = 0 Then
               b = Val(Right(txt, Len(txt) - pos2))
               Else
               b = Val(Mid(txt, pos2 + 1, pos3 - pos2))
               End If
            NColorSet(Nr) = RGB(R, G, b)
            CustPal.palPalEntry(Nr).peRed = Chr(R)
            CustPal.palPalEntry(Nr).peGreen = Chr(G)
            CustPal.palPalEntry(Nr).peBlue = Chr(b)
            Nr = Nr + 1
         Wend
      Close Ch
      If ColCount <> APICells Then
         MsgBox Str(APICells) & " colors extracted from a " & Str(ColCount) & "-colors PAL file."
         End If

   Case "MAP"
      Ch = FreeFile
      Nr = 0
      Open CD_File.FileName For Input As Ch
         While Not EOF(Ch)
            Line Input #Ch, txt
            If Nr < APICells Then
               pos1 = InStr(1, txt, " ")
               pos2 = InStr(pos1 + 1, txt, " ")
               pos3 = InStr(pos2 + 1, txt, " ")
               R = Val(Left(txt, pos1 - 1))
               G = Val(Mid(txt, pos1 + 1, pos2 - pos1))
               If pos3 = 0 Then
                  b = Val(Right(txt, Len(txt) - pos2))
                  Else
                  b = Val(Mid(txt, pos2 + 1, pos3 - pos2))
                  End If
               NColorSet(Nr) = RGB(R, G, b)
               CustPal.palPalEntry(Nr).peRed = Chr(R)
               CustPal.palPalEntry(Nr).peGreen = Chr(G)
               CustPal.palPalEntry(Nr).peBlue = Chr(b)
               End If
            Nr = Nr + 1
         Wend
      Close Ch
      ColCount = Nr - 1
      If ColCount <> APICells Then
         MsgBox Str(APICells) & " colors extracted from a" & Str(ColCount) & "-colors Fractint MAP file."
         End If

   Case "BMP"
      ColCount = PALfromBMP(CD_File.FileName)
      MsgBox Str(ColCount) & " colors extracted from bitmap"
      If ColCount = 0 Then Exit Sub
   End Select

   PalInPicPal
   ShowPallet

End Sub

Private Sub cmdOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim I As Long, mX As Long, mY As Long
   Dim Ret As Long, hPal As Long
   
   BaseColcnt = NBaseColcnt
   MakeFluent = chkFluent.Value
   For I = 1 To 16: UseColor(I) = NUseColor(I): Next I
   For I = 0 To 255: ColorSet(I) = NColorSet(I): Next I
   Clipboard.Clear
   Ret = OpenClipboard(hwnd)
   If Ret = 0 Then MsgBox "clipboard error": End
   hPal = CreatePalette(CustPal)
   SetClipboardData CF_PALETTE, hPal
   CloseClipboard
   frmPnt.picPal.Picture = Clipboard.GetData(CF_PALETTE)
   frmPnt.picEx.Picture = Clipboard.GetData(CF_PALETTE): frmPnt.picEx.PSet (0, 0), ColorSet(BGCol)
   StretchBlt frmPnt.picEx.hdc, 0, 0, SzeX, SzeY, frmPnt.picUndo.hdc, 0, 0, SzeX, SzeY, SRCCOPY
   frmPnt.picUndo.Picture = Clipboard.GetData(CF_PALETTE): frmPnt.picUndo.PSet (0, 0), ColorSet(BGCol)
   frmPnt.picEd.Picture = Clipboard.GetData(CF_PALETTE): frmPnt.picEd.PSet (0, 0), ColorSet(BGCol)
   frmPnt.picMse.Picture = Clipboard.GetData(CF_PALETTE): frmPnt.picMse.PSet (0, 0), ColorSet(BGCol)
   frmPnt.picPat.Picture = Clipboard.GetData(CF_PALETTE)
   DeleteObject hPal
   For I = 0 To 255
      mX = 1 + (I Mod 8) * 8
      mY = 1 + (I \ 8) * 8
      frmPnt.picPal.Line (mX, mY)-Step(6, 6), ColorSet(I), BF
   Next I
   OK = True: Unload Me
End Sub

Private Sub Form_Activate()
   Dim I As Long
   
   If Loading = True Then
      BColNr = 1
      PalInPicPal
      For I = 0 To 9
      If Val(cmbBaseColcnt.List(I)) = BaseColcnt Then cmbBaseColcnt.ListIndex = I
      Next I
      Loading = False
      End If
End Sub

Private Sub Form_Load()
   Dim I As Long
   Loading = True
   
'   BaseColorStart
   cmbBaseColcnt.AddItem "2"
   cmbBaseColcnt.AddItem "3"
   cmbBaseColcnt.AddItem "4"
   cmbBaseColcnt.AddItem "5"
   cmbBaseColcnt.AddItem "6"
   cmbBaseColcnt.AddItem "8"
   cmbBaseColcnt.AddItem "10"
   cmbBaseColcnt.AddItem "12"
   cmbBaseColcnt.AddItem "15"
   cmbBaseColcnt.AddItem "16"
   For I = 0 To 15: NUseColor(I) = UseColor(I): Next I
   For I = 0 To 255: NColorSet(I) = ColorSet(I): Next I
   For I = 1 To 15: Load linBCol(I): Load picBcol(I): Next I
   For I = 0 To 15: picBcol(I).BackColor = NUseColor(I): Next I
   chkFluent = MakeFluent
End Sub

Private Sub Form_Paint()
   BevelObject Me, 114, 4, 377, 172, 0
   BevelObject Me, 117, 8, 374, 38, 1
   BevelObject Me, 117, 42, 374, 104, 1
   BevelObject Me, 117, 108, 374, 168, 1
End Sub

Private Sub picBcol_click(Index As Integer)
   BColNr = Index
   shpSelect.Left = picBcol(BColNr).Left - 1
   shpSelect.Top = picBcol(BColNr).Top - 1
   shpSelect.Visible = True
   ShowRGB picBcol(BColNr).BackColor
End Sub

Private Sub picBcol_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)
   picBcol(Index).BackColor = DragColor
   picBcol_click Index
   'If Index = BColNr Then ShowRGB DragColor
   'cmdMakePal.Enabled = True
End Sub

Private Sub picBcol_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then DragColor = picBcol(Index).BackColor: picBcol(Index).Drag 1
End Sub

Private Sub picPal_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      DragColor = picPal.Point(X, Y)
      DragIndex = X
      picPal.Drag 1
      End If
End Sub

Private Sub hsRGB_Change(Index As Integer)
   ShowRGB RGB(hsRGB(0), hsRGB(1), hsRGB(2))
   cmdMakePal.Enabled = True
End Sub

Private Sub hsRGB_Scroll(Index As Integer)
   ShowRGB RGB(hsRGB(0), hsRGB(1), hsRGB(2))
   cmdMakePal.Enabled = True
End Sub

