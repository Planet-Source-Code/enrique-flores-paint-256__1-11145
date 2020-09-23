VERSION 5.00
Begin VB.Form frmInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Info..."
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3270
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2295
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      DownPicture     =   "frmInfo.frx":0000
      Height          =   540
      Left            =   2400
      Picture         =   "frmInfo.frx":03E6
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1605
      Width           =   840
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   480
      Left            =   165
      Picture         =   "frmInfo.frx":07CC
      Top             =   1305
      Width           =   480
   End
   Begin VB.Label Panel3D4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "98/08/2000"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   135
      TabIndex        =   4
      Top             =   1965
      Width           =   765
   End
   Begin VB.Label Panel3D3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "by"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1320
      TabIndex        =   3
      Top             =   1305
      Width           =   435
   End
   Begin VB.Label Panel3D2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Enrique A. Flores B. and Kew Lung"
      ForeColor       =   &H00FF0000&
      Height          =   675
      Left            =   915
      TabIndex        =   2
      Top             =   1560
      Width           =   1275
   End
   Begin VB.Label lblVersion 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "v1.01"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   285
      Left            =   2475
      TabIndex        =   1
      Top             =   1215
      Width           =   615
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   1005
      Left            =   135
      Picture         =   "frmInfo.frx":0AD6
      Stretch         =   -1  'True
      Top             =   135
      Width           =   2955
   End
End
Attribute VB_Name = "frmInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   lblVersion.Caption = "v" & App.Major & "." & Format(App.Revision, "00")
   Label1.ZOrder
End Sub

Private Sub Label1_Click()
   cmdOK_Click
End Sub

