VERSION 5.00
Begin VB.Form frmFill 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Colors - Fill Style"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ControlBox      =   0   'False
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2970
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   360
      Left            =   780
      TabIndex        =   13
      Top             =   900
      Width           =   3510
      Begin VB.OptionButton pshFld 
         Height          =   315
         Index           =   9
         Left            =   2835
         Picture         =   "frmFill.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshFld 
         Height          =   315
         Index           =   8
         Left            =   2520
         Picture         =   "frmFill.frx":0102
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshFld 
         Height          =   315
         Index           =   7
         Left            =   2205
         Picture         =   "frmFill.frx":0204
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshFld 
         Height          =   315
         Index           =   6
         Left            =   1890
         Picture         =   "frmFill.frx":0306
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshFld 
         Height          =   315
         Index           =   5
         Left            =   1575
         Picture         =   "frmFill.frx":0408
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshFld 
         Height          =   315
         Index           =   4
         Left            =   1260
         Picture         =   "frmFill.frx":050A
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshFld 
         Height          =   315
         Index           =   3
         Left            =   945
         Picture         =   "frmFill.frx":060C
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshFld 
         Height          =   315
         Index           =   2
         Left            =   630
         Picture         =   "frmFill.frx":070E
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshFld 
         Height          =   315
         Index           =   1
         Left            =   315
         Picture         =   "frmFill.frx":0810
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshFld 
         Height          =   315
         Index           =   0
         Left            =   0
         Picture         =   "frmFill.frx":0912
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
   End
   Begin VB.TextBox txtPSzeY 
      Height          =   300
      Left            =   4050
      TabIndex        =   11
      Text            =   "8"
      Top             =   1425
      Width           =   345
   End
   Begin VB.TextBox txtPSzeX 
      Height          =   300
      Left            =   3435
      TabIndex        =   10
      Text            =   "8"
      Top             =   1425
      Width           =   345
   End
   Begin VB.OptionButton pshFill 
      Height          =   315
      Index           =   4
      Left            =   105
      Picture         =   "frmFill.frx":0A14
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1890
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshFill 
      Height          =   315
      Index           =   3
      Left            =   105
      Picture         =   "frmFill.frx":0B0E
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1365
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshFill 
      Height          =   315
      Index           =   2
      Left            =   105
      Picture         =   "frmFill.frx":0C08
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   570
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshFill 
      Height          =   315
      Index           =   1
      Left            =   105
      Picture         =   "frmFill.frx":0D0A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   105
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   315
   End
   Begin VB.CommandButton cmdCancel 
      DownPicture     =   "frmFill.frx":0E04
      Height          =   540
      Left            =   2370
      Picture         =   "frmFill.frx":12A6
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2340
      Width           =   2190
   End
   Begin VB.CommandButton cmdOK 
      DownPicture     =   "frmFill.frx":1748
      Height          =   540
      Left            =   105
      Picture         =   "frmFill.frx":1B2E
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2340
      Width           =   2190
   End
   Begin VB.Label lbl 
      Caption         =   "x"
      Height          =   165
      Index           =   4
      Left            =   3870
      TabIndex        =   12
      Top             =   1470
      Width           =   180
   End
   Begin VB.Label lbl 
      Caption         =   "&Clipboard bitmap-pattern : tile image in clipboard"
      Height          =   195
      Index           =   3
      Left            =   495
      TabIndex        =   9
      Top             =   1935
      Width           =   3525
   End
   Begin VB.Label lbl 
      Caption         =   "&Bitmap-pattern : fluent colors grid - size"
      Height          =   195
      Index           =   2
      Left            =   495
      TabIndex        =   8
      Top             =   1425
      Width           =   2760
   End
   Begin VB.Label lbl 
      Caption         =   "&Fluent colors : from left mouse color to right"
      Height          =   195
      Index           =   1
      Left            =   495
      TabIndex        =   7
      Top             =   645
      Width           =   4110
   End
   Begin VB.Label lbl 
      Caption         =   "&Normal : color bound to left mouse button"
      Height          =   195
      Index           =   0
      Left            =   495
      TabIndex        =   6
      Top             =   165
      Width           =   3525
   End
End
Attribute VB_Name = "frmFill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdCancel.Value = True
   Hide
End Sub

Private Sub cmdOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim I As Long
   For I = 1 To 4
   If pshFill(I).Value = True Then
      FillType = I
      frmPnt.cmdSetFillStyle.Picture = pshFill(I).Picture
      Exit For
      End If
   Next I
   For I = 0 To 9
      If pshFld(I).Value = True Then FloodType = I: Exit For
   Next I
   PSzeX = Val(txtPSzeX.Text)
   PSzeY = Val(txtPSzeY.Text)
   cmdOK.Value = False
   Hide
End Sub

Private Sub Form_Activate()
   pshFill(FillType).Value = True
   pshFld(FloodType).Value = True
   txtPSzeX.Text = PSzeX
   txtPSzeY.Text = PSzeY
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 13: cmdOk_MouseUp 1, 0, 2, 2: KeyCode = 0
      Case 27: cmdCancel_MouseUp 1, 0, 2, 2: KeyCode = 0
      Case Asc("N"): pshFill(1).Value = True: KeyCode = 0
      Case Asc("F"): pshFill(2).Value = True: KeyCode = 0
      Case Asc("B"): pshFill(3).Value = True: KeyCode = 0
      Case Asc("C"): pshFill(4).Value = True: KeyCode = 0
      Case Asc("0"): pshFld(1).Value = True: KeyCode = 0
      Case Asc("1"): pshFld(4).Value = True: KeyCode = 0
      Case Asc("2"): pshFld(2).Value = True: KeyCode = 0
      Case Asc("3"): pshFld(5).Value = True: KeyCode = 0
      Case Asc("4"): pshFld(3).Value = True: KeyCode = 0
      Case Asc("5"): pshFld(6).Value = True: KeyCode = 0
      Case Asc("6"): pshFld(10).Value = True: KeyCode = 0
      Case Asc("7"): pshFld(9).Value = True: KeyCode = 0
      Case Asc("8"): pshFld(7).Value = True: KeyCode = 0
      Case Asc("9"): pshFld(8).Value = True: KeyCode = 0
   End Select
End Sub

Private Sub pshFld_Click(Index As Integer)
   If pshFill(1).Value = True Then pshFill(2).Value = True
End Sub

Private Sub txtPszeX_GotFocus()
   txtPSzeX.SelStart = 0
   txtPSzeX.SelLength = Len(txtPSzeX.Text)
End Sub

Private Sub txtPszeX_KeyPress(Keyascii As Integer)
   If Keyascii = 8 Then Exit Sub
   If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9")) Then Keyascii = 0
End Sub

Private Sub txtPszeY_GotFocus()
   txtPSzeY.SelStart = 0
   txtPSzeY.SelLength = Len(txtPSzeY.Text)
End Sub

Private Sub txtPszeY_KeyPress(Keyascii As Integer)
   If Keyascii = 8 Then Exit Sub
   If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9")) Then Keyascii = 0
End Sub

