VERSION 5.00
Begin VB.Form frmClpMd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Paste Clipboard Mode"
   ClientHeight    =   2430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2430
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox kadO 
      Height          =   2265
      Left            =   3510
      ScaleHeight     =   2205
      ScaleWidth      =   2325
      TabIndex        =   10
      Top             =   90
      Width           =   2385
      Begin VB.PictureBox kadI 
         BackColor       =   &H80000011&
         BorderStyle     =   0  'None
         Height          =   1590
         Left            =   0
         ScaleHeight     =   1590
         ScaleWidth      =   2025
         TabIndex        =   13
         Top             =   0
         Width           =   2025
         Begin VB.PictureBox pic 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1200
            Left            =   30
            ScaleHeight     =   80
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   115
            TabIndex        =   14
            Top             =   30
            Width           =   1725
         End
      End
      Begin VB.HScrollBar HS 
         Height          =   270
         Left            =   0
         TabIndex        =   12
         Top             =   1920
         Width           =   1050
      End
      Begin VB.VScrollBar VS 
         Height          =   915
         Left            =   2055
         TabIndex        =   11
         Top             =   15
         Width           =   270
      End
   End
   Begin VB.CommandButton cmdCancel 
      DownPicture     =   "frmClpMd.frx":0000
      Height          =   540
      Left            =   1725
      Picture         =   "frmClpMd.frx":04A2
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   1785
      Width           =   1485
   End
   Begin VB.CommandButton cmdOK 
      DownPicture     =   "frmClpMd.frx":0944
      Height          =   540
      Left            =   195
      Picture         =   "frmClpMd.frx":0D2A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   1785
      Width           =   1485
   End
   Begin VB.Frame fra 
      Height          =   1005
      Index           =   0
      Left            =   75
      TabIndex        =   4
      Top             =   15
      Width           =   3345
      Begin VB.OptionButton optadCut 
         Caption         =   "Don't adapt, &Cut where needed"
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   195
         Value           =   -1  'True
         Width           =   3135
      End
      Begin VB.OptionButton optadStretch 
         Caption         =   "Adapt &Image to Surface (stretch)"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   450
         Width           =   3135
      End
      Begin VB.OptionButton optadSurface 
         Caption         =   "Adapt &Surface to Image (new size)"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   705
         Width           =   3135
      End
   End
   Begin VB.Frame fra 
      Height          =   540
      Index           =   1
      Left            =   60
      TabIndex        =   0
      Top             =   1095
      Width           =   3360
      Begin VB.OptionButton optrtAll 
         Caption         =   "&All"
         Height          =   195
         Left            =   1395
         TabIndex        =   2
         Top             =   225
         Width           =   615
      End
      Begin VB.OptionButton optrtSel 
         Caption         =   "&Selection"
         Height          =   195
         Left            =   2160
         TabIndex        =   1
         Top             =   225
         Width           =   975
      End
      Begin VB.Label lbl 
         Caption         =   "Regarding to"
         Height          =   195
         Left            =   165
         TabIndex        =   3
         Top             =   210
         Width           =   960
      End
   End
End
Attribute VB_Name = "frmClpMd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   OK = False
   Unload Me
End Sub

Private Sub cmdOk_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case True
      Case optadCut: ClipMode = 0
      Case optadStretch: ClipMode = 1
      Case optadSurface: ClipMode = 2
   End Select
   Select Case True
      Case optrtAll: AdaptTo = 0
      Case optrtSel: AdaptTo = 1
   End Select
   OK = True
   Unload Me
End Sub

Private Sub Form_Activate()
   CheckScrolls pic, kadI, kadO, HS, VS
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   Select Case KeyCode
      Case 13: cmdOk_MouseUp 1, 0, 2, 2: KeyCode = 0 ' enter
      Case 27: cmdCancel_MouseUp 1, 0, 2, 2: KeyCode = 0 ' escape
   End Select
End Sub

Private Sub Form_Load()
   Select Case ClipMode
      Case 0: optadCut.Value = True
      Case 1: optadStretch.Value = True
      Case 2: optadSurface.Value = True
   End Select
   Select Case AdaptTo
      Case 0: optrtAll.Value = True
      Case 1: optrtSel.Value = True
   End Select
   On Error Resume Next
   pic.Picture = frmPnt.picBuff.Picture
   If Err <> 0 Then MsgBox Error$
End Sub

Private Sub HS_Change()
   pic.Left = -HS.Value
End Sub

Private Sub HS_Scroll()
   pic.Left = -HS.Value
End Sub

Private Sub VS_Change()
   pic.Top = -VS.Value
End Sub

Private Sub VS_Scroll()
   pic.Top = -VS.Value
End Sub

