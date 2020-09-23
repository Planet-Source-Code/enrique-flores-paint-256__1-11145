VERSION 5.00
Begin VB.Form frmSchad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Schades"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1860
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2265
   ScaleWidth      =   1860
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancel 
      DownPicture     =   "frmSchad.frx":0000
      Height          =   540
      Left            =   930
      Picture         =   "frmSchad.frx":04A2
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   1695
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      DownPicture     =   "frmSchad.frx":0944
      Height          =   540
      Left            =   15
      Picture         =   "frmSchad.frx":0D2A
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1695
      Width           =   885
   End
   Begin VB.CheckBox chkShDouble 
      Caption         =   "&Double"
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   135
      TabIndex        =   18
      Top             =   1395
      Width           =   870
   End
   Begin VB.OptionButton pshShC 
      Height          =   315
      Index           =   3
      Left            =   1455
      Picture         =   "frmSchad.frx":1110
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1260
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshShC 
      Height          =   315
      Index           =   2
      Left            =   1140
      Picture         =   "frmSchad.frx":120A
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1260
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshShC 
      Height          =   315
      Index           =   1
      Left            =   1455
      Picture         =   "frmSchad.frx":1304
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   930
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshShC 
      Height          =   315
      Index           =   0
      Left            =   1140
      Picture         =   "frmSchad.frx":13FE
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   930
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   315
   End
   Begin VB.TextBox txtShDist 
      Height          =   285
      Left            =   1125
      TabIndex        =   13
      Text            =   "1"
      Top             =   285
      Width           =   360
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   60
      TabIndex        =   3
      Top             =   300
      Width           =   945
      Begin VB.OptionButton pshShad 
         Height          =   315
         Index           =   8
         Left            =   630
         Picture         =   "frmSchad.frx":14F8
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   660
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshShad 
         Height          =   315
         Index           =   7
         Left            =   315
         Picture         =   "frmSchad.frx":15FA
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   660
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshShad 
         Height          =   315
         Index           =   6
         Left            =   0
         Picture         =   "frmSchad.frx":16FC
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   660
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshShad 
         Height          =   315
         Index           =   5
         Left            =   630
         Picture         =   "frmSchad.frx":17FE
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   315
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshShad 
         Height          =   315
         Index           =   4
         Left            =   0
         Picture         =   "frmSchad.frx":1900
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   330
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshShad 
         Height          =   315
         Index           =   3
         Left            =   630
         Picture         =   "frmSchad.frx":1A02
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshShad 
         Height          =   315
         Index           =   2
         Left            =   315
         Picture         =   "frmSchad.frx":1B04
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshShad 
         Height          =   315
         Index           =   1
         Left            =   0
         Picture         =   "frmSchad.frx":1C06
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshShad 
         Height          =   315
         Index           =   0
         Left            =   330
         Picture         =   "frmSchad.frx":1D08
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   315
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   315
      End
   End
   Begin VB.Label lbl 
      Caption         =   "Color"
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   2
      Left            =   1125
      TabIndex        =   2
      Top             =   720
      Width           =   600
   End
   Begin VB.Label lbl 
      Caption         =   "Distance"
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   1
      Left            =   1125
      TabIndex        =   1
      Top             =   60
      Width           =   690
   End
   Begin VB.Label lbl 
      Caption         =   "Type"
      ForeColor       =   &H00C00000&
      Height          =   210
      Index           =   0
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   600
   End
End
Attribute VB_Name = "frmSchad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
   OK = False: Hide
End Sub

Private Sub cmdOK_Click()
   Dim I As Long
   
   For I = 0 To 8
   If pshShad(I).Value = True Then
      Shade = I
      frmPnt.cmdSchades.Picture = pshShad(I).Picture
      Exit For
      End If
   Next I
   ShadeDis = Val(txtShDist.Text)
   ShadeDouble = chkShDouble.Value
   For I = 0 To 3
      If pshShC(I).Value = True Then ShadeCol = I
   Next I
   OK = True
   Hide
End Sub

Private Sub Form_Activate()
   pshShad(Shade).Value = True
   txtShDist.Text = ShadeDis
   pshShC(ShadeCol).Value = True
   chkShDouble.Value = ShadeDouble
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 13: cmdOK_Click: KeyCode = 0
      Case vbKeyEscape: cmdCancel_Click: KeyCode = 0
      Case Asc("D"): chkShDouble = IIf(chkShDouble = 1, 0, 1)
      Case Asc("A"): txtShDist.SetFocus: KeyCode = 0
      Case vbKeyNumpad5: If pshShad(0) = False Then pshShad(0) = True: KeyCode = 0
      Case vbKeyNumpad7: If pshShad(1) = False Then pshShad(1) = True: KeyCode = 0
      Case vbKeyNumpad8: If pshShad(2) = False Then pshShad(2) = True: KeyCode = 0
      Case vbKeyNumpad9: If pshShad(3) = False Then pshShad(3) = True: KeyCode = 0
      Case vbKeyNumpad4: If pshShad(4) = False Then pshShad(4) = True: KeyCode = 0
      Case vbKeyNumpad6: If pshShad(5) = False Then pshShad(5) = True: KeyCode = 0
      Case vbKeyNumpad1: If pshShad(6) = False Then pshShad(6) = True: KeyCode = 0
      Case vbKeyNumpad2: If pshShad(7) = False Then pshShad(7) = True: KeyCode = 0
      Case vbKeyNumpad3: If pshShad(8) = False Then pshShad(8) = True: KeyCode = 0
      Case Asc("W"): If pshShC(0) = False Then pshShC(0) = True: KeyCode = 0
      Case Asc("B"): If pshShC(1) = False Then pshShC(1) = True: KeyCode = 0
      Case Asc("L"): If pshShC(2) = False Then pshShC(2) = True: KeyCode = 0
      Case Asc("R"): If pshShC(3) = False Then pshShC(3) = True: KeyCode = 0
   End Select
End Sub

Private Sub txtShDist_GotFocus()
   txtShDist.SelStart = 0
   txtShDist.SelLength = Len(txtShDist.Text)
   Me.KeyPreview = False
End Sub

Private Sub txtShDist_KeyPress(Keyascii As Integer)
   If Keyascii = 8 Then Exit Sub
   If Not (Keyascii > Asc("0") And Keyascii <= Asc("9")) Then Keyascii = 0
End Sub

Private Sub txtShDist_LostFocus()
   Me.KeyPreview = True
End Sub


