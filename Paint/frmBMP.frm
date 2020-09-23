VERSION 5.00
Begin VB.Form frmBMP 
   Caption         =   "BMP"
   ClientHeight    =   5070
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6915
   Icon            =   "frmBMP.frx":0000
   LinkTopic       =   "Form2"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5070
   ScaleMode       =   0  'User
   ScaleWidth      =   6915
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1155
      Left            =   45
      TabIndex        =   16
      Top             =   3330
      Visible         =   0   'False
      Width           =   4395
      Begin VB.TextBox Text1 
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   765
         Left            =   2235
         MultiLine       =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   300
         Width           =   1995
      End
      Begin VB.Label lblFileName 
         Caption         =   "File :"
         ForeColor       =   &H80000011&
         Height          =   195
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   4365
      End
      Begin VB.Label lblFileDateTime 
         Caption         =   "00/00/00"
         ForeColor       =   &H00800080&
         Height          =   225
         Left            =   0
         TabIndex        =   20
         Top             =   585
         Width           =   1635
      End
      Begin VB.Label lblFileLen 
         Caption         =   "00000 Bytes"
         ForeColor       =   &H00C00000&
         Height          =   195
         Left            =   0
         TabIndex        =   19
         Top             =   285
         Width           =   1635
      End
      Begin VB.Image Image1 
         Height          =   255
         Left            =   0
         Picture         =   "frmBMP.frx":030A
         Top             =   855
         Width           =   240
      End
      Begin VB.Label lblSize 
         Caption         =   "0 x 0"
         Height          =   195
         Left            =   345
         TabIndex        =   18
         Top             =   900
         Width           =   1335
      End
   End
   Begin VB.PictureBox kadPalO 
      Height          =   405
      Left            =   45
      ScaleHeight     =   345
      ScaleWidth      =   6705
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4575
      Visible         =   0   'False
      Width           =   6765
      Begin VB.PictureBox kadPalI 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   330
         ScaleHeight     =   315
         ScaleWidth      =   3045
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   15
         Width           =   3045
         Begin VB.PictureBox picPal 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   210
            Left            =   75
            ScaleHeight     =   12
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   51
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   45
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdDown 
         Caption         =   ">"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3600
         TabIndex        =   13
         Top             =   15
         Width           =   315
      End
      Begin VB.CommandButton cmdUp 
         Caption         =   "<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   15
         TabIndex        =   12
         Top             =   30
         Width           =   300
      End
   End
   Begin VB.CommandButton cmdDelFile 
      Height          =   345
      Left            =   2295
      Picture         =   "frmBMP.frx":0414
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Delete File"
      Top             =   75
      UseMaskColor    =   -1  'True
      Width           =   420
   End
   Begin VB.CommandButton cmdExit 
      Height          =   345
      Left            =   6435
      Picture         =   "frmBMP.frx":050E
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Close window"
      Top             =   75
      UseMaskColor    =   -1  'True
      Width           =   420
   End
   Begin VB.CommandButton cmdCopy 
      Height          =   345
      Left            =   5955
      Picture         =   "frmBMP.frx":0608
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Copy image"
      Top             =   75
      UseMaskColor    =   -1  'True
      Width           =   420
   End
   Begin VB.PictureBox kadO 
      Height          =   2790
      Left            =   4500
      ScaleHeight     =   2730
      ScaleWidth      =   2295
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   480
      Width           =   2355
      Begin VB.VScrollBar VScroll 
         Height          =   915
         Left            =   1995
         TabIndex        =   7
         Top             =   15
         Width           =   270
      End
      Begin VB.HScrollBar HScroll 
         Height          =   270
         Left            =   0
         TabIndex        =   6
         Top             =   2460
         Width           =   1050
      End
      Begin VB.PictureBox kadI 
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Height          =   1590
         Left            =   0
         ScaleHeight     =   1590
         ScaleWidth      =   1875
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   1875
         Begin VB.PictureBox picBMP 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ClipControls    =   0   'False
            ForeColor       =   &H80000008&
            Height          =   1200
            Left            =   30
            ScaleHeight     =   80
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   97
            TabIndex        =   5
            Top             =   30
            Width           =   1455
         End
      End
   End
   Begin VB.FileListBox File1 
      Height          =   2820
      Left            =   2265
      Pattern         =   "*.bmp;*.ico;*.emf;*.gif"
      TabIndex        =   2
      Top             =   480
      Width           =   2145
   End
   Begin VB.DirListBox Dir1 
      Height          =   2790
      Left            =   45
      TabIndex        =   1
      Top             =   495
      Width           =   2145
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   75
      Width           =   2145
   End
End
Attribute VB_Name = "frmBMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim PicFExt As String      ' picture file extention
Dim ButtonDown As Boolean  ' commandbutton is down (cmdDown, cmdUp)
Dim FileName As String

Private Sub CopyFile()
   Dim ipb As String
   
   If File1.ListIndex <> -1 Then
      ipb = InputBox("Type a new name for the copy of '" & FileName & "'", "Copy File", FileName)
      If ipb = "" Then Exit Sub
      On Error Resume Next
      Screen.MousePointer = vbHourglass
      FileCopy FileName, ipb
      If Err <> 0 Then MsgBox Err.Description
      On Error GoTo 0
      File1.Refresh
      Screen.MousePointer = vbDefault
      End If
End Sub

' show extracted bitmap-file colors in picPal
Private Sub BMPcolors(ByRef Col As String, ByRef ColCount As Long)
   Dim D As Integer ' size of one color rectangle
   Dim R As Integer, G As Integer, b As Integer
   Dim I As Long, k As Long
   
   picPal.Left = 0
   Clipboard.Clear
   Clipboard.SetData picBMP.Image ', 9
   picPal.Picture = Clipboard.GetData(9)
   If ColCount <= 256 Then ' only < 16-bit mode
      D = picPal.Height \ 15
      picPal.Width = ColCount * D * 15
      For I = 1 To Len(Col) - 1 Step 4
        b = Asc(Mid(Col, I, 1))
        G = Asc(Mid(Col, I + 1, 1))
        R = Asc(Mid(Col, I + 2, 1))
        picPal.Line (k * D + 1, 1)-Step(D - 2, D - 2), QBColor(15), B
        picPal.Line (k * D, 0)-Step(D - 2, D - 2), 0, B
        picPal.Line (k * D + 1, 1)-Step(D - 4, D - 4), RGB(R, G, b), BF
        k = k + 1
      Next I
      End If
End Sub

Private Function BMPHeader21$(BI As BITMAPINFOHEADER)  '40 bytes
   Dim txt As String
   txt = txt & "Planes :" & Str(BI.biPlanes) & vbCrLf
   txt = txt & "BitCount :" & Str(BI.biBitCount) & vbCrLf
   txt = txt & "Compression :" & Str(BI.biCompression) & vbCrLf
   txt = txt & "ClrUsed :" & Str(BI.biClrUsed) & vbCrLf
   txt = txt & "ClrImportant :" & Str(BI.biClrImportant) & vbCrLf
   BMPHeader21$ = txt
End Function

' extract bitmap info from file, return in text format
Private Function BMPprop() As String
  Dim BFh As BITMAPFILEHEADER
  Dim BIh As BITMAPINFO          ' 40 bytes+colors
  Dim Ch As Long                 ' file channel
  Dim PalBytes As Long, ColCount As Long
  Dim txt As String, Pal As String
  
  Ch = FreeFile
  Open FileName For Binary As Ch
  Get #Ch, 1, BFh
  Get #Ch, , BIh
  txt = txt & BMPHeader21$(BIh.bmiHeader)
  PalBytes = BFh.bfOffBits - BIh.bmiHeader.biSize - 14
  If PalBytes > 0 Then
     Pal = Left(BIh.bmiColors, PalBytes)
     If BIh.bmiHeader.biClrUsed = 0 Then
        ColCount = 2 ^ BIh.bmiHeader.biBitCount
        Else
        ColCount = BIh.bmiHeader.biClrUsed
        End If
     BMPcolors Pal, ColCount
     kadPalO.Visible = True
     Else
     kadPalO.Visible = False ' no pallet bytes > 256-colors mode
     End If
  Close Ch

BMPprop = txt
End Function

Private Sub cmdDelFile_click()
   If File1.ListIndex <> -1 Then
      If MsgBox("Are you sure you want to delete  '" & FileName & "'  wilt wissen", vbYesNo, "Delete File") = vbNo Then Exit Sub
      On Error Resume Next
      Kill FileName
      If Err <> 0 Then MsgBox Err.Description
      On Error GoTo 0
      File1.Refresh
      File1_PathChange
      End If
End Sub

Private Sub cmdCopy_Click()
   Clipboard.Clear
   Clipboard.SetData picBMP.Image, 2
End Sub

Private Sub cmdDown_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Pz As Long
   
   ButtonDown = True: Pz = 600
   Do While ButtonDown = True
     If picPal.Left + picPal.Width > kadPalI.Width Then
        picPal.Left = picPal.Left - picPal.Height * 4
        Else
        Exit Do
        End If
     Pause Pz
     If Pz > 200 Then Pz = Pz / 2
   Loop
End Sub

Private Sub cmdDown_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ButtonDown = False
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdUp_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim Pz As Long
   
   ButtonDown = True: Pz = 600
   Do While ButtonDown = True
     If picPal.Left < 0 Then
        picPal.Left = picPal.Left + picPal.Height * 4
        Else
        Exit Do
        End If
     Pause Pz
     If Pz > 200 Then Pz = Pz / 2
   Loop

End Sub

Private Sub cmdUp_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   ButtonDown = False
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
   On Error Resume Next
   Dir1.Path = DrivePaths(Drive1.ListIndex)
   If Err <> 0 Then MsgBox Err.Description
End Sub

Private Sub Drive1_GotFocus()
   prevDrive = Drive1.ListIndex
   DrivePaths(prevDrive) = Dir1.Path
End Sub

' show selected image and it's properties
Private Sub File1_Click()
   Dim DT As Variant
   
   Screen.MousePointer = vbHourglass

   FileName = FixPath(Dir1.Path) & File1.FileName
   lblFileName.Caption = FileName
   lblFileLen.Caption = Format(FileLen(FileName), "###,#0") & " bytes"
   DT = FileDateTime(FileName)
   lblFileDateTime.Caption = Left(DT, InStr(DT, " "))
   PicFExt = UCase(Right(FileName, 3))
   
   On Error Resume Next
   picBMP.Picture = LoadPicture(FileName)
   
   If Err <> 0 Then
      MsgBox Err.Description
      Else
      CheckScrolls picBMP, kadI, kadO, HScroll, VScroll
      lblSize.Caption = Str(picBMP.Width \ 15) & " x" & Str(picBMP.Height \ 15)
      Select Case PicFExt
         Case "BMP": Text1.Text = BMPprop()
         Case "ICO": Text1.Text = ICOprop()
      End Select
      Frame1.Visible = True
      kadPalO.Visible = True
      End If
   On Error GoTo 0
   Screen.MousePointer = vbDefault
End Sub

Private Sub File1_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
   Case 45: CopyFile
   Case 46: cmdDelFile_click
   End Select
End Sub

Private Sub File1_PathChange()
   If File1.ListCount = 0 Then
      picBMP.Picture = LoadPicture()
      Frame1.Visible = False
      kadPalO.Visible = False
      Else
      File1.ListIndex = 0
      End If
End Sub

Private Sub Form_Load()
   Dim D As Long
   
   ReDim DrivePaths(Drive1.ListCount)
   For D = 0 To Drive1.ListCount - 1
     DrivePaths(D) = Left(Drive1.List(D), 2)
   Next D
   prevDrive = Drive1.ListIndex
   If prevPath = "" Then
      Dir1.Path = App.Path
      On Error Resume Next
      Dir1.Path = ".."
      On Error GoTo 0
      File1.Path = Dir1.Path
      Else
      Dir1.Path = prevPath
      End If
      
   picPal.Left = 0
End Sub

Private Sub Form_Resize()
   If Me.WindowState = 1 Then Exit Sub
   If Me.Width < 7035 Then Me.Width = 7035
   If Me.Height < 5595 Then Me.Height = 5595
   kadPalO.Width = Me.ScaleWidth - 2 * kadPalO.Left
   kadO.Width = Me.ScaleWidth - kadO.Left - kadPalO.Left
   kadPalO.Top = Me.ScaleHeight - kadPalO.Height - kadPalO.Left
   kadO.Height = Me.ScaleHeight - kadO.Top - kadPalO.Height - kadPalO.Left * 2
   cmdDown.Left = kadPalO.Width - cmdDown.Width - 60
   kadPalI.Width = kadPalO.Width - cmdDown.Width * 2 - 90
   CheckScrolls picBMP, kadI, kadO, HScroll, VScroll
End Sub

Private Sub Form_Unload(Cancel As Integer)
   prevPath = Dir1.Path
End Sub

Private Sub HScroll_Change()
   picBMP.Left = -HScroll.Value
End Sub

Private Sub HScroll_Scroll()
   picBMP.Left = -HScroll.Value
End Sub

' extract icon-file props and return them as a string
Private Function ICOprop() As String
   Dim txt As String, Pal As String
   Dim idReserved%
   Dim idType%
   Dim idCount%
   Dim Bte As Byte
   Dim wPlanes%
   Dim wBitCount%
   Dim dwBytesInRes As Long
   Dim dwImageOffset As Long
   Dim Ch As Long, PalBytes As Long, ColCount As Long
   
   On Error GoTo ICOpropFout:
   Ch = FreeFile
   Open FileName For Binary As Ch
   Get #Ch, , idReserved%
   Get #Ch, , idType%
   Get #Ch, , idCount%
   Get #Ch, , Bte
   Get #Ch, , Bte
   Get #Ch, , Bte
   Get #Ch, , Bte
   Get #Ch, , wPlanes%
   Get #Ch, , wBitCount%
   Get #Ch, , dwBytesInRes
   Get #Ch, , dwImageOffset
   
   Dim BIh As BITMAPINFO '40 bytes
   Get #Ch, dwImageOffset + 1, BIh.bmiHeader
   txt = txt & BMPHeader21$(BIh.bmiHeader)
   If BIh.bmiHeader.biSizeImage = 0 Then BIh.bmiHeader.biSizeImage = 640
   PalBytes = dwBytesInRes - BIh.bmiHeader.biSizeImage - BIh.bmiHeader.biSize
     
   If PalBytes > 0 Then
      Pal = String(PalBytes, " ")
      Get #Ch, , Pal
      BIh.bmiColors = Pal
      ColCount = 2 ^ BIh.bmiHeader.biBitCount
      BMPcolors Pal, ColCount
      kadPalO.Visible = True
      Else
      kadPalO.Visible = False
      End If
   
ICOpropEinde:
   ICOprop = txt
   Exit Function
ICOpropFout:
   MsgBox Err.Description
   Resume ICOpropEinde:
End Function

Private Sub VScroll_Change()
   picBMP.Top = -VScroll.Value
End Sub

Private Sub VScroll_Scroll()
   picBMP.Top = -VScroll.Value
End Sub

