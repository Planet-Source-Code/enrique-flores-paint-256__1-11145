VERSION 5.00
Begin VB.Form frmArrows 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Arrows"
   ClientHeight    =   1950
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3405
   ControlBox      =   0   'False
   FillStyle       =   0  'Solid
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   130
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   227
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.OptionButton optArr 
      Height          =   345
      Index           =   5
      Left            =   1905
      Picture         =   "frmArrows.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   150
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.OptionButton optArr 
      Height          =   345
      Index           =   4
      Left            =   1545
      Picture         =   "frmArrows.frx":00FA
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   150
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.OptionButton optArr 
      Height          =   345
      Index           =   3
      Left            =   1185
      Picture         =   "frmArrows.frx":01F4
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   150
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.OptionButton optArr 
      Height          =   345
      Index           =   2
      Left            =   825
      Picture         =   "frmArrows.frx":02EE
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   150
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.OptionButton optArr 
      Height          =   345
      Index           =   1
      Left            =   465
      Picture         =   "frmArrows.frx":03E8
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   150
      UseMaskColor    =   -1  'True
      Width           =   330
   End
   Begin VB.OptionButton optArr 
      Height          =   345
      Index           =   0
      Left            =   105
      Picture         =   "frmArrows.frx":04E2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   150
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   330
   End
   Begin VB.CommandButton cmdCancel 
      DownPicture     =   "frmArrows.frx":05DC
      Height          =   540
      Left            =   1725
      Picture         =   "frmArrows.frx":0A7E
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1305
      Width           =   1620
   End
   Begin VB.CommandButton cmdOK 
      DownPicture     =   "frmArrows.frx":0F20
      Height          =   540
      Left            =   60
      Picture         =   "frmArrows.frx":1306
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   1305
      Width           =   1590
   End
   Begin VB.CheckBox chkFilled 
      Caption         =   "&Filled"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2370
      TabIndex        =   6
      Top             =   225
      Value           =   1  'Checked
      Width           =   960
   End
   Begin VB.HScrollBar hsPar 
      Height          =   240
      Index           =   3
      LargeChange     =   2
      Left            =   1500
      Max             =   10
      TabIndex        =   4
      Top             =   900
      Value           =   10
      Width           =   1590
   End
   Begin VB.HScrollBar hsPar 
      Height          =   240
      Index           =   2
      LargeChange     =   2
      Left            =   1500
      Max             =   10
      TabIndex        =   1
      Top             =   615
      Value           =   10
      Width           =   1590
   End
   Begin VB.Label lblPar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "30"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   885
      TabIndex        =   5
      Top             =   900
      Width           =   600
   End
   Begin VB.Label lbl 
      Caption         =   "Head W"
      ForeColor       =   &H00800000&
      Height          =   180
      Index           =   3
      Left            =   135
      TabIndex        =   3
      Top             =   915
      Width           =   825
   End
   Begin VB.Label lblPar 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "10"
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   885
      TabIndex        =   2
      Top             =   615
      Width           =   600
   End
   Begin VB.Label lbl 
      Caption         =   "Width"
      ForeColor       =   &H00800000&
      Height          =   180
      Index           =   2
      Left            =   135
      TabIndex        =   0
      Top             =   630
      Width           =   825
   End
End
Attribute VB_Name = "frmArrows"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NArrowAngle As Long       ' local new values
Dim NArrowWidth As Long
Dim NArrowHeadW As Single
Dim NArrowDouble As Boolean
Dim NArrowFilled As Boolean

Private Sub chkFilled_Click()
   NArrowFilled = IIf(chkFilled.Value = 1, True, False)
End Sub

Private Sub cmdCancel_Click()
   OK = False
   Unload Me
End Sub

Private Sub cmdOK_Click()
   ArrowAngle = NArrowAngle   ' ok to set current values to new ones
   ArrowWidth = NArrowWidth
   ArrowHeadW = NArrowHeadW
   ArrowDouble = NArrowDouble
   ArrowFilled = NArrowFilled
   OK = True
   Unload Me
End Sub

Private Sub Form_Load()
   NArrowAngle = ArrowAngle ' pass cur. values to local new ones
   NArrowWidth = ArrowWidth
   NArrowHeadW = ArrowHeadW
   hsPar(2).Value = NArrowWidth
   hsPar(3).Value = NArrowHeadW
   NArrowDouble = ArrowDouble
   NArrowFilled = ArrowFilled
   chkFilled.Value = IIf(NArrowFilled = True, 1, 0)
   optArr(ArrowIndex).Value = True
End Sub

Private Sub hsPar_Change(Index As Integer)
   Select Case Index
      Case 2: NArrowWidth = hsPar(Index).Value
      Case 3: NArrowHeadW = hsPar(Index).Value
   End Select
   lblPar(Index).Caption = hsPar(Index).Value
   If Index = 0 Then lblPar(Index).Caption = lblPar(Index).Caption & "°"
End Sub

Private Sub hsPar_Scroll(Index As Integer)
   Select Case Index
      Case 2: NArrowWidth = hsPar(Index).Value
      Case 3: NArrowHeadW = hsPar(Index).Value
   End Select
   lblPar(Index).Caption = hsPar(Index).Value
   If Index = 0 Then lblPar(Index).Caption = lblPar(Index).Caption & "°"
End Sub

Private Sub optArr_Click(Index As Integer)
   Select Case Index
      Case 0: NArrowDouble = False: NArrowAngle = 90
      Case 1: NArrowDouble = False: NArrowAngle = 270
      Case 2: NArrowDouble = False: NArrowAngle = 180
      Case 3: NArrowDouble = False: NArrowAngle = 0
      Case 4: NArrowDouble = True: NArrowAngle = 0
      Case 5: NArrowDouble = True: NArrowAngle = 90
   End Select
   ArrowIndex = Index
End Sub

