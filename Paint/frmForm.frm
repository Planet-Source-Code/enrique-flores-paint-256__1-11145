VERSION 5.00
Begin VB.Form frmForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Star Formula"
   ClientHeight    =   4365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4365
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtAngle 
      Height          =   285
      Left            =   2475
      TabIndex        =   15
      Text            =   "270"
      Top             =   3330
      Width           =   435
   End
   Begin VB.HScrollBar HS 
      Height          =   255
      Index           =   2
      Left            =   90
      Max             =   360
      TabIndex        =   14
      Top             =   3345
      Value           =   270
      Width           =   2355
   End
   Begin VB.TextBox txtGPlus 
      Height          =   285
      Left            =   2475
      TabIndex        =   12
      Text            =   "120"
      Top             =   2790
      Width           =   435
   End
   Begin VB.HScrollBar HS 
      Height          =   255
      Index           =   1
      Left            =   90
      Max             =   180
      TabIndex        =   11
      Top             =   2805
      Value           =   120
      Width           =   2355
   End
   Begin VB.CommandButton cmdPaste 
      Enabled         =   0   'False
      Height          =   345
      Left            =   3030
      Picture         =   "frmForm.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Waarden uit klembord halen"
      Top             =   3000
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdCopy 
      Height          =   330
      Left            =   3030
      Picture         =   "frmForm.frx":0142
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Huidige waarden in klembord plaatsen"
      Top             =   2640
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.TextBox txtPts 
      Height          =   285
      Left            =   2475
      TabIndex        =   7
      Text            =   "4"
      Top             =   2265
      Width           =   435
   End
   Begin VB.HScrollBar HS 
      Height          =   255
      Index           =   0
      Left            =   90
      Max             =   360
      TabIndex        =   6
      Top             =   2280
      Value           =   4
      Width           =   2355
   End
   Begin VB.PictureBox pic 
      BackColor       =   &H8000000E&
      Height          =   1545
      Left            =   1875
      ScaleHeight     =   1485
      ScaleWidth      =   1470
      TabIndex        =   4
      Top             =   420
      Width           =   1530
   End
   Begin VB.CheckBox chkFilled 
      Caption         =   "&Filled"
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   1890
      TabIndex        =   3
      Top             =   135
      Width           =   1485
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   90
      TabIndex        =   2
      Top             =   135
      Width           =   1635
   End
   Begin VB.CommandButton cmdOK 
      DownPicture     =   "frmForm.frx":023C
      Height          =   540
      Left            =   105
      Picture         =   "frmForm.frx":0622
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3750
      Width           =   1590
   End
   Begin VB.CommandButton cmdCancel 
      DownPicture     =   "frmForm.frx":0A08
      Height          =   540
      Left            =   1770
      Picture         =   "frmForm.frx":0EAA
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3750
      Width           =   1620
   End
   Begin VB.Label lbl 
      Caption         =   "Start &Agle"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   13
      Top             =   3120
      Width           =   1620
   End
   Begin VB.Label lbl 
      Caption         =   "Angle &Step"
      ForeColor       =   &H00C00000&
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   10
      Top             =   2580
      Width           =   1620
   End
   Begin VB.Label lbl 
      Caption         =   "Number of &Points"
      ForeColor       =   &H00C00000&
      Height          =   225
      Index           =   0
      Left            =   90
      TabIndex        =   5
      Top             =   2055
      Width           =   1620
   End
End
Attribute VB_Name = "frmForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Pts As Long, GPlus As Long, Angle As Long
Dim pPts As Long, pGPlus As Long, pAngle As Long ' previous

Private Sub ShowFormula()
   pic.Cls
   Pts = Val(txtPts)
   GPlus = Val(txtGPlus)
   Angle = Val(txtAngle)
   DrawFormula pic, 4, 4, 92, 92, Pts, GPlus, Angle, ColorSet(0), chkFilled.Value
   HS(0).Value = Pts
   HS(1).Value = GPlus
   HS(2).Value = Angle
End Sub

Private Sub chkFilled_Click()
   ShowFormula
End Sub

Private Sub cmdCancel_Click()
   OK = False: Hide
End Sub

Private Sub cmdCopy_Click()
   pPts = Val(txtPts.Text)
   pGPlus = Val(txtGPlus.Text)
   pAngle = Val(txtAngle.Text)
   cmdPaste.Enabled = True
End Sub

Private Sub cmdOK_Click()
   FormulaPts = Val(frmForm.txtPts)
   FormulaGplus = Val(frmForm.txtGPlus)
   FormulaAngle = Val(frmForm.txtAngle)
   FormulaFilled = chkFilled.Value
   OK = True
   Hide
End Sub

Private Sub cmdPaste_click()
   txtPts.Text = pPts
   txtGPlus.Text = pGPlus
   txtAngle.Text = pAngle
   ShowFormula
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 13: cmdOK_Click: KeyCode = 0
      Case 27: cmdCancel_Click: KeyCode = 0
      Case Asc("F"): KeyCode = 0
          If chkFilled.Value = True Then
             chkFilled.Value = False
             Else
             chkFilled.Value = True
             End If
      Case Asc("P"): HS(0).SetFocus: KeyCode = 0
      Case Asc("S"): HS(1).SetFocus: KeyCode = 0
      Case Asc("A"): HS(2).SetFocus: KeyCode = 0
   End Select
End Sub

Private Sub Form_Load()
   List1.AddItem "Five-Star"
   List1.AddItem "Eight-Star"
   List1.AddItem "Twelve-Star"
   List1.AddItem "Eight-Star2"
   List1.AddItem "Triangle"
   List1.AddItem "Rectangle"
   List1.AddItem "Pentagon"
   List1.AddItem "Hexagon"
   List1.AddItem "Eight-Angle"
   List1.AddItem "Twelve-Angle"
   List1.ListIndex = 0
End Sub

Private Sub Form_Paint()
   ShowFormula
End Sub

Private Sub HS_Change(Index As Integer)
   Select Case Index
      Case 0: txtPts.Text = HS(0).Value
      Case 1: txtGPlus.Text = HS(1).Value
      Case 2: txtAngle.Text = HS(2).Value
   End Select
   ShowFormula
End Sub

Private Sub HS_Scroll(Index As Integer)
   Select Case Index
      Case 0: txtPts.Text = HS(0).Value
      Case 1: txtGPlus.Text = HS(1).Value
      Case 2: txtAngle.Text = HS(2).Value
   End Select
   ShowFormula
End Sub

Private Sub List1_Click()
   pic.Cls
   Select Case List1.List(List1.ListIndex)
      Case "Five-Star": Pts = 6: GPlus = 144
      Case "Eight-Star": Pts = 9: GPlus = 135
      Case "Twelve-Star": Pts = 13: GPlus = 150
      Case "Eight-Star2": Pts = 10: GPlus = 160
      Case "Triangle": Pts = 4: GPlus = 120
      Case "Rectangle": Pts = 5: GPlus = 90
      Case "Pentagon": Pts = 6: GPlus = 72
      Case "Hexagon": Pts = 7: GPlus = 60
      Case "Eight-Angle": Pts = 9: GPlus = 45
      Case "Twelve-Angle": Pts = 13: GPlus = 30
   End Select
   Angle = 270
   txtPts = Pts
   txtGPlus = GPlus
   txtAngle = Angle
   ShowFormula
End Sub

Private Sub txtPts_Change()
   If Val(txtPts) > HS(0).Max Then txtPts = HS(0).Max
   If Val(txtPts) < HS(0).Min Then txtPts = HS(0).Min
End Sub

Private Sub txtPts_GotFocus()
   txtPts.SelStart = 0
   txtPts.SelLength = Len(txtPts.Text)
End Sub

Private Sub txtGplus_Change()
   If Val(txtGPlus) > HS(1).Max Then txtGPlus = HS(1).Max
   If Val(txtGPlus) < HS(1).Min Then txtGPlus = HS(1).Min
End Sub

Private Sub txtGplus_GotFocus()
   txtGPlus.SelStart = 0
   txtGPlus.SelLength = Len(txtGPlus.Text)
End Sub

Private Sub txtAngle_Change()
   If Val(txtAngle) > HS(2).Max Then txtAngle = HS(2).Max
   If Val(txtAngle) < HS(2).Min Then txtAngle = HS(2).Min
End Sub

Private Sub txtAngle_GotFocus()
   txtAngle.SelStart = 0
   txtAngle.SelLength = Len(txtAngle.Text)
End Sub

