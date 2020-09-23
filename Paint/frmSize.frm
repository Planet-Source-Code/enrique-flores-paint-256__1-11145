VERSION 5.00
Begin VB.Form frmSize 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Size"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3255
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2235
   ScaleWidth      =   3255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chkShowGrid 
      Caption         =   "Show grid"
      ForeColor       =   &H00800000&
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   1275
      Value           =   1  'Checked
      Width           =   1500
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      DownPicture     =   "frmSize.frx":0000
      Height          =   540
      Left            =   1650
      Picture         =   "frmSize.frx":04A2
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1590
      Width           =   1485
   End
   Begin VB.CommandButton cmdOK 
      Default         =   -1  'True
      DownPicture     =   "frmSize.frx":0944
      Height          =   540
      Left            =   120
      Picture         =   "frmSize.frx":0D2A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   1590
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Caption         =   "Scale factor"
      ForeColor       =   &H00800000&
      Height          =   675
      Left            =   75
      TabIndex        =   4
      Top             =   480
      Width           =   3090
      Begin VB.OptionButton optZoom 
         Caption         =   "x8"
         Height          =   255
         Index           =   4
         Left            =   2400
         TabIndex        =   11
         Top             =   255
         Value           =   -1  'True
         Width           =   510
      End
      Begin VB.OptionButton optZoom 
         Caption         =   "x6"
         Height          =   255
         Index           =   3
         Left            =   1851
         TabIndex        =   10
         Top             =   255
         Width           =   510
      End
      Begin VB.OptionButton optZoom 
         Caption         =   "x4"
         Height          =   255
         Index           =   2
         Left            =   1304
         TabIndex        =   9
         Top             =   255
         Width           =   510
      End
      Begin VB.OptionButton optZoom 
         Caption         =   "x2"
         Height          =   255
         Index           =   1
         Left            =   757
         TabIndex        =   8
         Top             =   255
         Width           =   510
      End
      Begin VB.OptionButton optZoom 
         Caption         =   "x1"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   5
         Top             =   255
         Width           =   510
      End
   End
   Begin VB.TextBox txtHeight 
      Height          =   285
      Left            =   2250
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   75
      Width           =   780
   End
   Begin VB.TextBox txtWidth 
      Height          =   285
      Left            =   735
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   75
      Width           =   780
   End
   Begin VB.Label lbl 
      Caption         =   "&Height"
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   1
      Left            =   1635
      TabIndex        =   2
      Top             =   120
      Width           =   600
   End
   Begin VB.Label lbl 
      Caption         =   "&Width"
      ForeColor       =   &H00800000&
      Height          =   210
      Index           =   0
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   585
   End
End
Attribute VB_Name = "frmSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NPixSze As Long ' size of a PixSze according to scale

' restrict to workable values
Private Sub CheckSize()
   Dim V1 As Long, V2 As Long
   
   V1 = Val(txtWidth.Text)
   V2 = Val(txtHeight.Text)
   If V1 > 273 Or V2 > 273 Then
      optZoom(4).Enabled = False
      If NPixSze > 6 Then NPixSze = 6
      Else
      optZoom(4).Enabled = True
      End If
   If V1 > 364 Or V2 > 364 Then
      optZoom(3).Enabled = False
      If NPixSze > 4 Then NPixSze = 4
      Else
      optZoom(3).Enabled = True
      End If
   If V1 > 546 Or V2 > 546 Then
      optZoom(2).Enabled = False
      If NPixSze > 2 Then NPixSze = 2
      Else
      optZoom(2).Enabled = True
      End If
   If V1 > MaxSize Or V2 > MaxSize Then
      If NPixSze > 2 Then NPixSze = 2
      If V1 > MaxSize Then txtWidth.Text = Format(MaxSize)
      If V2 > MaxSize Then txtHeight.Text = Format(MaxSize)
      End If
   Select Case NPixSze
      Case 1: optZoom(0).Value = True
      Case 2: optZoom(1).Value = True
      Case 4: optZoom(2).Value = True
      Case 6: optZoom(3).Value = True
      Case 8: optZoom(4).Value = True
   End Select
End Sub

Private Sub cmdCancel_Click()
   OK = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
   SzeX = Val(txtWidth)
   SzeY = Val(txtHeight)
   If SzeX < 4 Then SzeX = 4
   If SzeY < 4 Then SzeY = 4
   PixSze = NPixSze
   ShowGrid = IIf(chkShowGrid.Value = 1, True, False)
   OK = True
   Unload Me
End Sub

Private Sub Form_Load()
   NPixSze = PixSze
   txtWidth.Text = SzeX
   txtHeight.Text = SzeY
   chkShowGrid.Value = IIf(ShowGrid = True, 1, 0)
End Sub

Private Sub optZoom_click(Index As Integer)
   Select Case Index
      Case 0: NPixSze = 1: chkShowGrid.Value = 0: chkShowGrid.Enabled = False
      Case 1: NPixSze = 2: chkShowGrid.Value = 0: chkShowGrid.Enabled = False
      Case 2: NPixSze = 4: chkShowGrid.Enabled = True
      Case 3: NPixSze = 6: chkShowGrid.Enabled = True
      Case 4: NPixSze = 8: chkShowGrid.Enabled = True
   End Select
End Sub

Private Sub txtWidth_Change()
   CheckSize
End Sub

Private Sub txtWidth_GotFocus()
   txtWidth.SelStart = 0
   txtWidth.SelLength = Len(txtWidth.Text)
End Sub

Private Sub txtWidth_KeyPress(Keyascii As Integer)
   If Keyascii = 8 Then Exit Sub
   If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9")) Then Keyascii = 0
End Sub

Private Sub txtHeight_Change()
   CheckSize
End Sub

Private Sub txtHeight_GotFocus()
   txtHeight.SelStart = 0
   txtHeight.SelLength = Len(txtHeight.Text)
End Sub

Private Sub txtHeight_KeyPress(Keyascii As Integer)
   If Keyascii = 8 Then Exit Sub
   If Not (Keyascii >= Asc("0") And Keyascii <= Asc("9")) Then Keyascii = 0
End Sub

