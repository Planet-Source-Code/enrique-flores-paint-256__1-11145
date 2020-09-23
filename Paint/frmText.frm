VERSION 5.00
Begin VB.Form frmText 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Text"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3945
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   3945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox kadAlign 
      Height          =   405
      Left            =   75
      ScaleHeight     =   345
      ScaleWidth      =   1005
      TabIndex        =   10
      Top             =   75
      Width           =   1065
      Begin VB.OptionButton pshAlign 
         Height          =   315
         Index           =   0
         Left            =   15
         Picture         =   "frmText.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   15
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshAlign 
         Height          =   315
         Index           =   1
         Left            =   345
         Picture         =   "frmText.frx":014E
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   15
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshAlign 
         Height          =   315
         Index           =   2
         Left            =   675
         Picture         =   "frmText.frx":029C
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   15
         UseMaskColor    =   -1  'True
         Width           =   315
      End
   End
   Begin VB.PictureBox KadHk 
      Height          =   675
      Left            =   225
      ScaleHeight     =   615
      ScaleWidth      =   1185
      TabIndex        =   5
      Top             =   75
      Width           =   1245
      Begin VB.OptionButton pshAng 
         Height          =   270
         Index           =   0
         Left            =   30
         Picture         =   "frmText.frx":03EA
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   30
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   555
      End
      Begin VB.OptionButton pshAng 
         Height          =   555
         Index           =   1
         Left            =   600
         Picture         =   "frmText.frx":04C0
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   285
      End
      Begin VB.OptionButton pshAng 
         Height          =   555
         Index           =   2
         Left            =   900
         Picture         =   "frmText.frx":059A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   30
         UseMaskColor    =   -1  'True
         Width           =   285
      End
      Begin VB.OptionButton pshAng 
         Height          =   270
         Index           =   3
         Left            =   30
         Picture         =   "frmText.frx":0674
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   315
         UseMaskColor    =   -1  'True
         Width           =   555
      End
   End
   Begin VB.CommandButton cmdFont 
      Appearance      =   0  'Flat
      Caption         =   "&Font"
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   60
      Width           =   2325
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   45
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmText.frx":0756
      Top             =   780
      Width           =   3825
   End
   Begin VB.PictureBox picEx 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   60
      ScaleHeight     =   125
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   250
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1935
      Width           =   3810
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   -240
         Top             =   120
         Width           =   195
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      DownPicture     =   "frmText.frx":0760
      Height          =   540
      Left            =   2010
      Picture         =   "frmText.frx":0C02
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4005
      Width           =   1860
   End
   Begin VB.CommandButton cmdOK 
      DownPicture     =   "frmText.frx":10A4
      Height          =   540
      Left            =   75
      Picture         =   "frmText.frx":148A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4005
      Width           =   1830
   End
   Begin VB.Label lbldX 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   30
      TabIndex        =   17
      Top             =   1725
      Width           =   405
   End
   Begin VB.Label lbldY 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   615
      TabIndex        =   16
      Top             =   1725
      Width           =   405
   End
   Begin VB.Label Lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "x"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   450
      TabIndex        =   15
      Top             =   1725
      Width           =   135
   End
   Begin VB.Label lblLett 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "LblLett"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   1590
      TabIndex        =   14
      Top             =   525
      Width           =   2325
   End
End
Attribute VB_Name = "frmText"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TA As Long, TH As Long ' text align/angle(dutch:Hoek)

Private Sub FontDialog()
   Dim fntFlags As DlgFontFlags
   
   fntFlags = CF_BOTH Or CF_PRINTERFONTS Or CF_TTONLY
   CD_Font.FontName = picEx.FontName
   CD_Font.FontSize = picEx.FontSize
   CD_Font.FontBold = picEx.FontBold
   CD_Font.FontItalic = picEx.FontItalic
   CD_Font.DialogTitle = "Font Instellen"
   CD_Font.Flags = fntFlags              ' true type printer fonts only
   On Error Resume Next
   CD_Font.ShowFont                      ' activate Font dialog
   If Err <> 0 Then Exit Sub
   On Error GoTo 0
   picEx.FontName = CD_Font.FontName
   picEx.FontSize = CD_Font.FontSize
   picEx.FontBold = CD_Font.FontBold
   picEx.FontItalic = CD_Font.FontItalic
End Sub

Private Sub cmdCancel_Click()
   OK = False
   Unload Me
End Sub

Private Sub cmdFont_click()
   Dim txt As String
   
   FontDialog
   txt = picEx.FontName & Str(picEx.FontSize)
   If picEx.FontBold = True Then txt = txt & " B"
   If picEx.FontItalic = True Then txt = txt & " I"
   lblLett = txt
   ShowText
   picEx.SetFocus
End Sub

Private Sub cmdOK_Click()
   Text = Text1.Text
   TextFntNm = picEx.FontName
   TextSize = picEx.FontSize
   TextBold = picEx.FontBold
   TextItalic = picEx.FontItalic
   TextAlign = TA
   TextAngle = TH
   OK = True
   Unload Me
End Sub

Private Sub Form_Activate()
   ShowText
   Text1.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 13: cmdOK_Click: KeyCode = 0
      Case 27: cmdCancel_Click: KeyCode = 0
      Case Asc("F"): cmdFont_click: KeyCode = 0
   End Select
End Sub

Private Sub Form_Load()
   Dim txt As String
   
   Text1.Text = Text
   Select Case TextAlign
      Case 0: pshAlign(0).Value = True
      Case 1: pshAlign(2).Value = True
      Case 2: pshAlign(1).Value = True
   End Select
   Select Case TextAngle
      Case 0: pshAng(0).Value = True
      Case 900: pshAng(1).Value = True
      Case -900: pshAng(2).Value = True
      Case 1800: pshAng(3).Value = True
   End Select
   picEx.FontName = TextFntNm
   picEx.FontSize = TextSize
   picEx.FontBold = TextBold
   picEx.FontItalic = TextItalic
   txt = picEx.FontName & Str(picEx.FontSize)
   If picEx.FontBold = True Then txt = txt & " B"
   If picEx.FontItalic = True Then txt = txt & " I"
   lblLett = txt
   
   KadHk.Left = kadAlign.Left
   
End Sub

Private Sub picEx_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim X1 As Long, Y1 As Long, X2 As Long, Y2 As Long
   X1 = X - Abs(TextdX) / 2
   Y1 = Y - Abs(TextdY) / 2
   X2 = X + Abs(TextdX) / 2
   Y2 = Y + Abs(TextdY) / 2
   Shape1.Left = X1
   Shape1.Top = Y1
   Shape1.Width = Abs(TextdX)
   Shape1.Height = Abs(TextdY)
   APIText picEx, Int(X), Int(Y), Text1.Text, TA, TH, QBColor(12)
End Sub

Private Sub pshAlign_Click(Index As Integer)
   TA = Index
   ShowText
End Sub

Private Sub pshAng_Click(Index As Integer)
   Select Case Index
      Case 0: TH = 0
      Case 1: TH = 900
      Case 2: TH = -900
      Case 3: TH = 1800
   End Select
   ShowText
End Sub

Private Sub Text1_Change()
   ShowText
End Sub

Private Sub ShowText()
   Dim mX As Integer, mY As Integer
   Dim ct$
   Dim H As Long, Wt As Long, I As Long, E As Long
   Dim W, O, U, S, C, OP, CP, Q, PAF
   Dim F As String
   Dim hFnt As Long, hFntOld As Long
   Dim tdX As Long, tdY As Long
   Dim mX1 As Long, mY1 As Long


   mX = picEx.ScaleWidth \ 2
   mY = picEx.ScaleHeight \ 2
   picEx.Cls
   ct$ = Text1.Text
   
   If InStr(ct$, vbCrLf) Then
      KadHk.Visible = False
      kadAlign.Visible = True
      RC.Left = 0: RC.Top = 0
      DrawText picEx.hdc, ct$, Len(ct$), RC, TA Or DT_CALCRECT
      OffsetRect RC, mX - RC.Right \ 2, mY - RC.Bottom \ 2
      RC.Left = RC.Left - 6
      RC.Right = RC.Right + 6
      DrawText picEx.hdc, ct$, Len(ct$), RC, TA
      picEx.Line (RC.Left, RC.Top)-(RC.Right, RC.Bottom), QBColor(8), B
      tdX = RC.Right - RC.Left + 1
      tdY = RC.Bottom - RC.Top + 1
      Else
      KadHk.Visible = True
      kadAlign.Visible = False
      Dim TM As TEXTMETRIC, Sz As Size
      GetTextMetrics picEx.hdc, TM
      H = TM.tmHeight
      Wt = TM.tmWeight
      I = TM.tmItalic 'Asc(TM.tmItalic)
      F$ = String$(128, " ")
      GetTextFace picEx.hdc, 128, F$
      E = TH
      hFnt = CreateFont(H, W, E, O, Wt, I, U, S, C, OP, CP, Q, PAF, F$)
      hFntOld = SelectObject(picEx.hdc, hFnt)
      GetTextExtentPoint picEx.hdc, ct$, Len(ct$), Sz
      Select Case True
      Case pshAng(0)
         tdX = Sz.cX: tdY = Sz.cY
         mX1 = mX - tdX \ 2: mY1 = mY - tdY \ 2
      Case pshAng(1)
         tdX = Sz.cY: tdY = Sz.cX
         mX1 = mX - tdX \ 2: mY1 = mY + tdY \ 2: tdY = -tdY
      Case pshAng(2)
         tdX = Sz.cY: tdY = Sz.cX
         mX1 = mX + tdX \ 2: tdX = -tdX: mY1 = mY - tdY \ 2
      Case pshAng(3)
         tdX = Sz.cX: tdY = Sz.cY
         mX1 = mX + tdX \ 2: tdX = -tdX: mY1 = mY + tdY \ 2: tdY = -tdY
      End Select
      TextOut picEx.hdc, mX1, mY1, ct, Len(ct)
      SelectObject picEx.hdc, hFntOld
      DeleteObject hFnt
      picEx.Line (mX1, mY1)-Step(tdX, tdY), QBColor(8), B
      End If
   'label2.Caption = mX1
   'label3.Caption = mY1
   lblDX.Caption = Abs(tdX)
   lblDY.Caption = Abs(tdY)
   TextdX = tdX
   TextdY = tdY
End Sub

Private Sub Text1_GotFocus()
   Me.KeyPreview = False
End Sub

Private Sub Text1_LostFocus()
   Me.KeyPreview = True
End Sub

