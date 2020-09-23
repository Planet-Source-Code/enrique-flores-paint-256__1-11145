VERSION 5.00
Begin VB.Form frmPnt 
   Caption         =   "Paint256 by: Enrique A. Flores B. and Kew Lung"
   ClientHeight    =   5820
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6960
   Icon            =   "frmPnt.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   5820
   ScaleWidth      =   6960
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picBG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1155
      ScaleHeight     =   225
      ScaleWidth      =   240
      TabIndex        =   82
      ToolTipText     =   "Background color ( click + control key )"
      Top             =   3885
      Width           =   270
   End
   Begin VB.OptionButton pshAct 
      Height          =   300
      Index           =   9
      Left            =   1140
      Picture         =   "frmPnt.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   81
      TabStop         =   0   'False
      ToolTipText     =   "Arrows *"
      Top             =   3465
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.PictureBox picBuff 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   2205
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   80
      Top             =   5310
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picPat 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   15
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   79
      Top             =   6360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.PictureBox picUndo 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   2925
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   78
      Top             =   5325
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3570
      Top             =   5115
   End
   Begin VB.PictureBox StatusBar 
      ForeColor       =   &H000000C0&
      Height          =   360
      Left            =   30
      ScaleHeight     =   20
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   456
      TabIndex        =   60
      Top             =   5400
      Width           =   6900
      Begin VB.Label lblDY 
         Caption         =   "0000"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   6315
         TabIndex        =   76
         Top             =   45
         Width           =   420
      End
      Begin VB.Label lbl 
         Caption         =   "dY:"
         Height          =   195
         Index           =   7
         Left            =   6045
         TabIndex        =   75
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblDX 
         Caption         =   "0000"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   5565
         TabIndex        =   74
         Top             =   45
         Width           =   420
      End
      Begin VB.Label lbl 
         Caption         =   "dX:"
         Height          =   195
         Index           =   6
         Left            =   5295
         TabIndex        =   73
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblY2 
         Caption         =   "0000"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   4635
         TabIndex        =   72
         Top             =   45
         Width           =   420
      End
      Begin VB.Label lbl 
         Caption         =   "Y2:"
         Height          =   195
         Index           =   5
         Left            =   4350
         TabIndex        =   71
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblX2 
         Caption         =   "0000"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   3795
         TabIndex        =   70
         Top             =   45
         Width           =   420
      End
      Begin VB.Label lbl 
         Caption         =   "X2:"
         Height          =   195
         Index           =   4
         Left            =   3510
         TabIndex        =   69
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblY1 
         Caption         =   "0000"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   2775
         TabIndex        =   68
         Top             =   45
         Width           =   420
      End
      Begin VB.Label lbl 
         Caption         =   "Y1:"
         Height          =   195
         Index           =   3
         Left            =   2490
         TabIndex        =   67
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblX1 
         Caption         =   "0000"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   1950
         TabIndex        =   66
         Top             =   45
         Width           =   420
      End
      Begin VB.Label lbl 
         Caption         =   "X1:"
         Height          =   195
         Index           =   2
         Left            =   1665
         TabIndex        =   65
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblY 
         Caption         =   "0000"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   960
         TabIndex        =   64
         Top             =   45
         Width           =   375
      End
      Begin VB.Label lbl 
         Caption         =   "Y:"
         Height          =   195
         Index           =   1
         Left            =   780
         TabIndex        =   63
         Top             =   45
         Width           =   240
      End
      Begin VB.Label lblX 
         Caption         =   "0000"
         ForeColor       =   &H00C00000&
         Height          =   210
         Left            =   255
         TabIndex        =   62
         Top             =   45
         Width           =   390
      End
      Begin VB.Label lbl 
         Caption         =   "X:"
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   61
         Top             =   45
         Width           =   150
      End
   End
   Begin VB.PictureBox kadWidths 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   1455
      ScaleHeight     =   390
      ScaleWidth      =   1680
      TabIndex        =   54
      Top             =   4590
      Visible         =   0   'False
      Width           =   1710
      Begin VB.OptionButton pshWidth 
         Height          =   300
         Index           =   4
         Left            =   1320
         Picture         =   "frmPnt.frx":0404
         Style           =   1  'Graphical
         TabIndex        =   59
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshWidth 
         Height          =   300
         Index           =   3
         Left            =   990
         Picture         =   "frmPnt.frx":04FE
         Style           =   1  'Graphical
         TabIndex        =   58
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshWidth 
         Height          =   300
         Index           =   2
         Left            =   675
         Picture         =   "frmPnt.frx":05F8
         Style           =   1  'Graphical
         TabIndex        =   57
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshWidth 
         Height          =   300
         Index           =   1
         Left            =   360
         Picture         =   "frmPnt.frx":06F2
         Style           =   1  'Graphical
         TabIndex        =   56
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshWidth 
         Height          =   300
         Index           =   0
         Left            =   30
         Picture         =   "frmPnt.frx":07EC
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   45
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   315
      End
   End
   Begin VB.PictureBox kadSh 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   1305
      ScaleHeight     =   705
      ScaleWidth      =   720
      TabIndex        =   49
      Top             =   3270
      Visible         =   0   'False
      Width           =   750
      Begin VB.OptionButton pshSh 
         Height          =   300
         Index           =   3
         Left            =   360
         Picture         =   "frmPnt.frx":08E6
         Style           =   1  'Graphical
         TabIndex        =   53
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshSh 
         Height          =   300
         Index           =   2
         Left            =   45
         Picture         =   "frmPnt.frx":09E0
         Style           =   1  'Graphical
         TabIndex        =   52
         Top             =   360
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshSh 
         Height          =   300
         Index           =   1
         Left            =   360
         Picture         =   "frmPnt.frx":0ADA
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   45
         UseMaskColor    =   -1  'True
         Width           =   315
      End
      Begin VB.OptionButton pshSh 
         Height          =   300
         Index           =   0
         Left            =   45
         Picture         =   "frmPnt.frx":0BD4
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   45
         UseMaskColor    =   -1  'True
         Value           =   -1  'True
         Width           =   315
      End
   End
   Begin VB.PictureBox kadRHCK 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   1470
      ScaleHeight     =   315
      ScaleWidth      =   2385
      TabIndex        =   46
      Top             =   1650
      Visible         =   0   'False
      Width           =   2415
      Begin VB.PictureBox picRounding 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   315
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   89
         TabIndex        =   48
         Top             =   45
         Width           =   1365
      End
      Begin VB.PictureBox picRHCK 
         BorderStyle     =   0  'None
         Height          =   195
         Left            =   2100
         ScaleHeight     =   13
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   15
         TabIndex        =   47
         Top             =   60
         Width           =   225
      End
      Begin VB.Image Image2 
         Height          =   225
         Left            =   45
         Picture         =   "frmPnt.frx":0CCE
         Top             =   45
         Width           =   225
      End
      Begin VB.Image Image1 
         Height          =   225
         Left            =   1725
         Picture         =   "frmPnt.frx":0DC8
         Top             =   45
         Width           =   225
      End
   End
   Begin VB.PictureBox kadEx 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   615
      Left            =   1545
      ScaleHeight     =   585
      ScaleWidth      =   945
      TabIndex        =   44
      Top             =   435
      Visible         =   0   'False
      Width           =   975
      Begin VB.PictureBox picEx 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   345
         Left            =   60
         ScaleHeight     =   23
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   34
         TabIndex        =   45
         Top             =   75
         Width           =   510
      End
   End
   Begin VB.PictureBox kadO 
      Height          =   2070
      Left            =   1560
      ScaleHeight     =   2010
      ScaleWidth      =   2400
      TabIndex        =   39
      Top             =   450
      Width           =   2460
      Begin VB.PictureBox kadEd 
         BackColor       =   &H80000011&
         BorderStyle     =   0  'None
         Height          =   1590
         Left            =   0
         ScaleHeight     =   1590
         ScaleWidth      =   2025
         TabIndex        =   42
         Top             =   0
         Width           =   2025
         Begin VB.PictureBox picEd 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1200
            Left            =   30
            ScaleHeight     =   80
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   115
            TabIndex        =   43
            Top             =   30
            Width           =   1725
            Begin VB.Line Line1 
               BorderStyle     =   3  'Dot
               Visible         =   0   'False
               X1              =   72
               X2              =   105
               Y1              =   31
               Y2              =   55
            End
            Begin VB.Shape shpBox 
               BorderStyle     =   3  'Dot
               Height          =   240
               Left            =   120
               Top             =   120
               Visible         =   0   'False
               Width           =   240
            End
         End
      End
      Begin VB.HScrollBar HSEd 
         Height          =   270
         Left            =   0
         TabIndex        =   41
         Top             =   1725
         Width           =   1050
      End
      Begin VB.VScrollBar VSEd 
         Height          =   915
         Left            =   2130
         TabIndex        =   40
         Top             =   15
         Width           =   270
      End
   End
   Begin VB.CommandButton cmdSchades 
      Height          =   300
      Left            =   1140
      Picture         =   "frmPnt.frx":0EC2
      Style           =   1  'Graphical
      TabIndex        =   38
      ToolTipText     =   "Set Shades"
      Top             =   4290
      Width           =   300
   End
   Begin VB.CommandButton cmdSetFillStyle 
      Height          =   300
      Left            =   1140
      Picture         =   "frmPnt.frx":0FC4
      Style           =   1  'Graphical
      TabIndex        =   37
      ToolTipText     =   "Set Fill Style"
      Top             =   5010
      Width           =   300
   End
   Begin VB.CommandButton cmdDrawWidth 
      Height          =   300
      Left            =   1140
      Picture         =   "frmPnt.frx":10BE
      Style           =   1  'Graphical
      TabIndex        =   36
      ToolTipText     =   "Set DrawWidth"
      Top             =   4650
      Width           =   300
   End
   Begin VB.OptionButton pshAct 
      Height          =   300
      Index           =   8
      Left            =   1140
      Picture         =   "frmPnt.frx":11B8
      Style           =   1  'Graphical
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Schades frame *"
      Top             =   3135
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshAct 
      Height          =   300
      Index           =   7
      Left            =   1140
      Picture         =   "frmPnt.frx":12B2
      Style           =   1  'Graphical
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Text *"
      Top             =   2805
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshAct 
      Height          =   300
      Index           =   6
      Left            =   1140
      Picture         =   "frmPnt.frx":13B4
      Style           =   1  'Graphical
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Star - formula *"
      Top             =   2475
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshAct 
      Height          =   315
      Index           =   5
      Left            =   1140
      Picture         =   "frmPnt.frx":14AE
      Style           =   1  'Graphical
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Fill"
      Top             =   2145
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshAct 
      Height          =   315
      Index           =   4
      Left            =   1140
      Picture         =   "frmPnt.frx":15A8
      Style           =   1  'Graphical
      TabIndex        =   31
      TabStop         =   0   'False
      ToolTipText     =   "Filled square - circle *"
      Top             =   1815
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshAct 
      Height          =   315
      Index           =   3
      Left            =   1140
      Picture         =   "frmPnt.frx":16A2
      Style           =   1  'Graphical
      TabIndex        =   30
      TabStop         =   0   'False
      ToolTipText     =   "Square - circle *"
      Top             =   1485
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshAct 
      Height          =   315
      Index           =   2
      Left            =   1140
      Picture         =   "frmPnt.frx":179C
      Style           =   1  'Graphical
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Lines"
      Top             =   1155
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshAct 
      Height          =   315
      Index           =   1
      Left            =   1140
      Picture         =   "frmPnt.frx":1896
      Style           =   1  'Graphical
      TabIndex        =   28
      TabStop         =   0   'False
      ToolTipText     =   "Points - freehand"
      Top             =   825
      UseMaskColor    =   -1  'True
      Value           =   -1  'True
      Width           =   315
   End
   Begin VB.OptionButton pshAct 
      Height          =   315
      Index           =   0
      Left            =   1140
      Picture         =   "frmPnt.frx":1990
      Style           =   1  'Graphical
      TabIndex        =   27
      TabStop         =   0   'False
      ToolTipText     =   "Select"
      Top             =   495
      UseMaskColor    =   -1  'True
      Width           =   315
   End
   Begin VB.PictureBox kad 
      Height          =   4875
      Left            =   30
      ScaleHeight     =   4815
      ScaleWidth      =   990
      TabIndex        =   22
      Top             =   450
      Width           =   1050
      Begin VB.PictureBox picPal 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   3840
         Left            =   15
         ScaleHeight     =   256
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   65
         TabIndex        =   26
         Top             =   930
         Width           =   975
         Begin VB.Shape shpMseR 
            BorderColor     =   &H00800080&
            BorderWidth     =   2
            Height          =   135
            Left            =   780
            Top             =   90
            Width           =   135
         End
         Begin VB.Shape shpMseL 
            BorderColor     =   &H00008000&
            BorderWidth     =   2
            Height          =   135
            Left            =   0
            Top             =   0
            Width           =   135
         End
      End
      Begin VB.PictureBox picMse 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   48
            Charset         =   2
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   885
         Left            =   45
         ScaleHeight     =   59
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   62
         TabIndex        =   23
         Top             =   45
         Width           =   930
         Begin VB.Label lblPalIDR 
            Caption         =   "007"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   660
            TabIndex        =   25
            Top             =   15
            Width           =   255
         End
         Begin VB.Label lblPalIDL 
            Caption         =   "000"
            BeginProperty Font 
               Name            =   "Small Fonts"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Left            =   15
            TabIndex        =   24
            Top             =   15
            Width           =   255
         End
      End
   End
   Begin VB.CommandButton cmdInfo 
      Height          =   345
      Left            =   6195
      Picture         =   "frmPnt.frx":1A8A
      Style           =   1  'Graphical
      TabIndex        =   16
      ToolTipText     =   "Info..."
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdHelp 
      Height          =   345
      Left            =   5835
      Picture         =   "frmPnt.frx":1B84
      Style           =   1  'Graphical
      TabIndex        =   15
      ToolTipText     =   "Help text"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdClear 
      Height          =   345
      Left            =   5370
      Picture         =   "frmPnt.frx":1C7E
      Style           =   1  'Graphical
      TabIndex        =   14
      ToolTipText     =   "Clear all or selection (bgcolor)"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdMirrV 
      Height          =   345
      Left            =   5010
      Picture         =   "frmPnt.frx":1D78
      Style           =   1  'Graphical
      TabIndex        =   13
      ToolTipText     =   "Mirror vertical all or selection"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdMirrH 
      Height          =   345
      Left            =   4650
      Picture         =   "frmPnt.frx":1E72
      Style           =   1  'Graphical
      TabIndex        =   12
      ToolTipText     =   "Mirror horizontal all or selection"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdMove 
      Height          =   345
      Left            =   4290
      Picture         =   "frmPnt.frx":1F6C
      Style           =   1  'Graphical
      TabIndex        =   11
      ToolTipText     =   "Move all or selection"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdPalShift 
      Height          =   345
      Left            =   3930
      Picture         =   "frmPnt.frx":2066
      Style           =   1  'Graphical
      TabIndex        =   10
      ToolTipText     =   "Pallete shift"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdPallet 
      Height          =   345
      Left            =   3510
      Picture         =   "frmPnt.frx":2160
      Style           =   1  'Graphical
      TabIndex        =   9
      ToolTipText     =   "Set 256 Pallete"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdSize 
      Height          =   345
      Left            =   3135
      Picture         =   "frmPnt.frx":225A
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Set Size & Scale"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdVB 
      Caption         =   "x1"
      Height          =   345
      Left            =   2775
      TabIndex        =   7
      ToolTipText     =   "Show real size"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdUndo 
      Height          =   345
      Left            =   2325
      Picture         =   "frmPnt.frx":2364
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Undo"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdPaste 
      Height          =   345
      Left            =   1890
      Picture         =   "frmPnt.frx":2466
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "Paste all or selection"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdCopy 
      Height          =   345
      Left            =   1530
      Picture         =   "frmPnt.frx":25A8
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Copy all or selection"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdPictures 
      Height          =   345
      Left            =   1170
      Picture         =   "frmPnt.frx":26A2
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Picture browser"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdOpen 
      Height          =   345
      Left            =   810
      Picture         =   "frmPnt.frx":279C
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Open picture"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdSave 
      Height          =   345
      Left            =   450
      Picture         =   "frmPnt.frx":289E
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Save picture"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.CommandButton cmdExit 
      Height          =   345
      Left            =   45
      Picture         =   "frmPnt.frx":2998
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Exit Paint256"
      Top             =   45
      UseMaskColor    =   -1  'True
      Width           =   360
   End
   Begin VB.PictureBox kadMove 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1170
      Left            =   4065
      ScaleHeight     =   1140
      ScaleWidth      =   795
      TabIndex        =   17
      Top             =   45
      Visible         =   0   'False
      Width           =   825
      Begin VB.CommandButton cmdMve 
         Height          =   315
         Index           =   3
         Left            =   240
         Picture         =   "frmPnt.frx":2A92
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   765
         Width           =   330
      End
      Begin VB.CommandButton cmdMve 
         Height          =   315
         Index           =   2
         Left            =   240
         Picture         =   "frmPnt.frx":2B8C
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   45
         Width           =   330
      End
      Begin VB.CommandButton cmdMve 
         Height          =   315
         Index           =   1
         Left            =   420
         Picture         =   "frmPnt.frx":2C8E
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   405
         Width           =   330
      End
      Begin VB.CommandButton cmdMve 
         Height          =   315
         Index           =   0
         Left            =   45
         Picture         =   "frmPnt.frx":2D90
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   405
         Width           =   330
      End
   End
   Begin VB.PictureBox picMask 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   570
      Left            =   15
      ScaleHeight     =   38
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   41
      TabIndex        =   77
      Top             =   5355
      Visible         =   0   'False
      Width           =   615
   End
End
Attribute VB_Name = "frmPnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' whenever you encounter the prefix 'kad',
' know it is a picturebox acting as a frame

'  p i c t u r e b o x e s
' picBG        background
' picBuff      buffer - real size
' picEd        edit - scaled
' picEx        example - image in real size
' picMask      mask
' picMse       mouse - shows selected left & right mouse colors
' picPal       pallet - to pick a color from with left & right mouse
' picPat       holds flood pattern
' picRHCK      shows selected rectangle/circle rounding
' picRounding  user points rounding % here
' picUndo      holds last image, ready to undo
' StatusBar    shows coordinates / progress

Dim Loading As Boolean              ' form is loading
Dim ColorL As Long                  ' color-index under mouse buttons
Dim ColorR As Long
Dim CurCol As Long                  ' current color = pallet ID
Dim BevelType As Long               ' bevel-frames

Dim Action As Long                  ' current Action type
Dim X1 As Long, Y1 As Long          ' anglepoints of limiting rectangle
Dim X2 As Long, Y2 As Long
Dim dX As Long, dY As Long          ' size of limiting rectangle

Dim mX1 As Long, mY1 As Long        ' limiting rectangle dragged with mouse
Dim mX2 As Long, mY2 As Long        ' or made by actions, or used for movements
Dim mX As Long, mY As Long          ' last or current mouseposition
Dim ButtonDown As Boolean           ' mousebutton is down
Dim MovingBox As Boolean            ' inside prev. rectangle-> moving box
Dim NewBox As Boolean               ' outside prev. rectangle-> new box

Dim prSzeX As Long                  ' previous size bmp
Dim prSzeY As Long
Dim pAction As Long                 ' previous action

' before any change, prepare undo
Private Sub BeforeUndo()
   StretchBlt picUndo.hdc, 0, 0, SzeX, SzeY, picEx.hdc, 0, 0, SzeX, SzeY, SRCCOPY
   picUndo.Refresh
   prSzeX = SzeX
   prSzeY = SzeY
   pAction = Action
End Sub

' draw the shades for the current action
' it's for every type of action completely the same
Private Sub DoActionShade()
   Dim Col As Long
   Dim Fl As Boolean
   Dim TOLEFT As Long, TORIGHT As Long
   Dim UPWARD  As Long, DOWNWARD  As Long
   
   Select Case ShadeCol
      Case 0: Col = 15
      Case 1: Col = 0
      Case 2: Col = ColorL
      Case 3: Col = ColorR
   End Select
   TOLEFT = 0: TORIGHT = 1: UPWARD = 2: DOWNWARD = 3

DoActionShadeDouble:
   Select Case Shade       ' first rotate
      Case 1: RotateAll TORIGHT, ShadeDis
              RotateAll DOWNWARD, ShadeDis
      Case 2: RotateAll DOWNWARD, ShadeDis
      Case 3: RotateAll TOLEFT, ShadeDis
              RotateAll DOWNWARD, ShadeDis
      Case 4: RotateAll TORIGHT, ShadeDis
      Case 5: RotateAll TOLEFT, ShadeDis
      Case 6: RotateAll TORIGHT, ShadeDis
              RotateAll UPWARD, ShadeDis
      Case 7: RotateAll UPWARD, ShadeDis
      Case 8: RotateAll TOLEFT, ShadeDis
              RotateAll UPWARD, ShadeDis
   End Select
   picEx.Refresh
   'DrawShade Col           ' then the actual drawing of the shade
   picPat.Width = SzeX * 15
   picPat.Height = SzeY * 15
   APIrect picPat.hdc, 0, Col, Col, 0, 0, SzeX, SzeY
   BitBlt picEx.hdc, 0, 0, SzeX, SzeY, picPat.hdc, 0, 0, SRCINVERT
   BitBlt picEx.hdc, 0, 0, SzeX, SzeY, picMask.hdc, 0, 0, SRCAND
   BitBlt picEx.hdc, 0, 0, SzeX, SzeY, picPat.hdc, 0, 0, SRCINVERT
   picEx.Refresh
   Select Case Shade       ' last rotate back to prev. position
      Case 1: RotateAll TOLEFT, ShadeDis
              RotateAll UPWARD, ShadeDis
      Case 2: RotateAll UPWARD, ShadeDis
      Case 3: RotateAll TORIGHT, ShadeDis
              RotateAll UPWARD, ShadeDis
      Case 4: RotateAll TOLEFT, ShadeDis
      Case 5: RotateAll TORIGHT, ShadeDis
      Case 6: RotateAll TOLEFT, ShadeDis
              RotateAll DOWNWARD, ShadeDis
      Case 7: RotateAll DOWNWARD, ShadeDis
      Case 8: RotateAll TORIGHT, ShadeDis
              RotateAll DOWNWARD, ShadeDis
   End Select
   picEx.Refresh
   If ShadeDouble = 1 And Fl = 0 Then
      Select Case ShadeCol ' opposite color
         Case 0: Col = 0
         Case 1: Col = 15
         Case 2: Col = ColorR
         Case 3: Col = ColorL
      End Select           ' opposite directions
      TOLEFT = 1: TORIGHT = 0: UPWARD = 3: DOWNWARD = 2
      Fl = 1               ' only do this ones of course
      GoTo DoActionShadeDouble
      End If
      
End Sub

' draws a beveled frame (raised, inset,...)
Private Sub DoBevel(pic As Control, ShT As Long)
   Dim kl1 As Long, kl2 As Long
   
   kl1 = 8: kl2 = 15
   Select Case ShT
   Case 0
   APIline pic.hdc, 0, 0, kl1, X1 + 1, Y1, X2 - 1, Y1
   APIline pic.hdc, 0, 0, kl1, X1, Y1 + 1, X1, Y2 - 1
   APIline pic.hdc, 0, 0, kl2, X1 + 1, Y2 - 1, X2 - 1, Y2 - 1
   APIline pic.hdc, 0, 0, kl2, X2 - 1, Y1 + 1, X2 - 1, Y2 - 1
   Case 1
   APIline pic.hdc, 0, 0, kl1, X1 + 1, Y1, X2 - 1, Y1
   APIline pic.hdc, 0, 0, kl1, X1, Y1 + 1, X1, Y2 - 1
   APIline pic.hdc, 0, 0, kl2, X1 + 1, Y2 - 1, X2 - 1, Y2 - 1
   APIline pic.hdc, 0, 0, kl2, X2 - 1, Y1 + 1, X2 - 1, Y2 - 1
   APIline pic.hdc, 0, 0, kl1, X1 + 2, Y1 + 1, X2 - 2, Y1 + 1
   APIline pic.hdc, 0, 0, kl1, X1 + 1, Y1 + 2, X1 + 1, Y2 - 2
   APIline pic.hdc, 0, 0, kl2, X1 + 2, Y2 - 2, X2 - 2, Y2 - 2
   APIline pic.hdc, 0, 0, kl2, X2 - 2, Y1 + 2, X2 - 2, Y2 - 2
   Case 2
   APIline pic.hdc, 0, 0, kl2, X1 + 1, Y1, X2 - 1, Y1
   APIline pic.hdc, 0, 0, kl2, X1, Y1 + 1, X1, Y2 - 1
   APIline pic.hdc, 0, 0, kl1, X1 + 1, Y2 - 1, X2 - 1, Y2 - 1
   APIline pic.hdc, 0, 0, kl1, X2 - 1, Y1 + 1, X2 - 1, Y2 - 1
   Case 3
   APIline pic.hdc, 0, 0, kl2, X1 + 1, Y1, X2 - 1, Y1
   APIline pic.hdc, 0, 0, kl2, X1, Y1 + 1, X1, Y2 - 1
   APIline pic.hdc, 0, 0, kl1, X1 + 1, Y2 - 1, X2 - 1, Y2 - 1
   APIline pic.hdc, 0, 0, kl1, X2 - 1, Y1 + 1, X2 - 1, Y2 - 1
   APIline pic.hdc, 0, 0, kl2, X1 + 2, Y1 + 1, X2 - 2, Y1 + 1
   APIline pic.hdc, 0, 0, kl2, X1 + 1, Y1 + 2, X1 + 1, Y2 - 2
   APIline pic.hdc, 0, 0, kl1, X1 + 2, Y2 - 2, X2 - 2, Y2 - 2
   APIline pic.hdc, 0, 0, kl1, X2 - 2, Y1 + 2, X2 - 2, Y2 - 2
   End Select
End Sub

' called from picEd
' shows the users drawings immediately but temporary
' real results follows later when mouse is released
' prepares mask too
Private Sub DrawPixel(X As Long, Y As Long, Color As Long)
   Dim eX1 As Long, eY1 As Long
   Dim eX2 As Long, eY2 As Long
   
   If PenWidth = 1 Then
      eX1 = X * PixSze
      eY1 = Y * PixSze
      eX2 = (X + 1) * PixSze
      eY2 = (Y + 1) * PixSze
      APIrect picEd.hdc, 0, 8, Color, eX1, eY1, eX2, eY2
      APIline picEx.hdc, 0, PenWidth, Color, X, Y, X + 1, Y + 1
      APIline picMask.hdc, 0, PenWidth, 0, X, Y, X + 1, Y + 1
      Else
      eX1 = (X - PenWidth \ 2) * PixSze
      eY1 = (Y - PenWidth \ 2) * PixSze
      If PenWidth = 3 Or PenWidth = 5 Then
         eX2 = (X + PenWidth \ 2 + 1) * PixSze
         eY2 = (Y + PenWidth \ 2 + 1) * PixSze
         Else
         eX2 = (X + PenWidth \ 2) * PixSze
         eY2 = (Y + PenWidth \ 2) * PixSze
         End If
      APIrect picEd.hdc, 0, Color, Color, eX1, eY1, eX2, eY2
      APIline picEx.hdc, 0, PenWidth, Color, X, Y, X, Y
      APIline picMask.hdc, 0, PenWidth, 0, X, Y, X, Y
      End If
   
   picEx.Refresh
   picEd.Refresh
   picMask.Refresh
End Sub

' is a black & white mask
Private Sub DrawMask()
   Dim Col As Long, ArrowRadii As Long
   
   If Action > A_PEN Then picMask.Cls ' mask is already made for A_PEN
   Select Case Action
   Case A_LINE
      APIline picMask.hdc, 0, PenWidth, 0, Line1.X1, Line1.Y1, Line1.X2, Line1.Y2
      picMask.Refresh
   Case A_RRECT, A_FRRECT
      DrawRRect picMask.hdc, 0
      picMask.Refresh
   Case A_FILL
      ' this is special, since you have to detect the filling-region
      ' to reproduce it in the mask
      ' starting with a blanco mask, first copy the All content inversed
      BitBlt picMask.hdc, 0, 0, SzeX, SzeY, picEx.hdc, 0, 0, SRCINVERT
      picMask.Refresh
      picEx.FillStyle = 0
      Col = picEx.Point(mX, mY) ' determine the color
      If CurCol <> 7 Then picEx.FillColor = 0 Else picEx.FillColor = ColorSet(CurCol)
      ' do the filling on the inversed image
      ExtFloodFill picEx.hdc, mX, mY, Col Xor &H2000000, 1 ' use nearest pallet index
      picEx.FillStyle = 1
      picEx.Refresh
      ' copy the All content inversed again so only the filled region remains
      BitBlt picMask.hdc, 0, 0, SzeX, SzeY, picEx.hdc, 0, 0, SRCINVERT
      picMask.Refresh
      picMask.FillStyle = 0
      picMask.FillColor = 0
      Col = picMask.Point(mX, mY)
      ' make the region black
      ExtFloodFill picMask.hdc, mX, mY, Col& Xor &H2000000, 1
      picMask.Refresh
      picMask.FillStyle = 1
   Case A_FORMULA
      DrawFormula picMask, X1, Y1, dX, dY, FormulaPts, FormulaGplus, FormulaAngle, 0, FormulaFilled
   Case A_TEXT
      picMask.FontName = TextFntNm: picMask.FontSize = TextSize
      picMask.FontBold = TextBold: picMask.FontItalic = TextItalic
      APIText picMask, X1 + dX \ 2, Y1 + dY \ 2, Text, TextAlign, TextAngle, 0
      picMask.Refresh
   Case A_ARROW
      picMask.FillStyle = IIf(ArrowFilled = True, 0, 1)
      ArrowRadii = IIf(dX <= dY, dY, dX) / 2
      DrawArrow picMask.hdc, X1 + dX \ 2, Y1 + dY \ 2, ArrowAngle, dX \ 2, dY \ 2, ArrowWidth, ArrowHeadW, ArrowDouble
      picMask.Refresh
      picMask.FillStyle = 1

   End Select
End Sub

' showing left & right mouse colors
Private Sub DrawMouse()
   picMse.Cls
   picMse.ForeColor = QBColor(8)
   picMse.CurrentX = -6
   picMse.CurrentY = -2
   picMse.Print Chr(56)
   picMse.ForeColor = QBColor(15)
   picMse.CurrentX = -7
   picMse.CurrentY = -3
   picMse.Print Chr(56)
   picMse.FillColor = ColorSet(ColorL)
   picMse.Line (26, 18)-(34, 23), ColorSet(ColorL), BF
   picMse.FillColor = ColorSet(ColorR)
   picMse.Line (39, 18)-(43, 23), ColorSet(ColorR), BF
   picMse.Refresh
End Sub

' drawing many grid lines slows down the proces
Private Sub DrawGrid()
   Dim I As Long
   
   For I = 0 To SzeY - 1
      APIline picEd.hdc, 0, 0, 8, 0, I * PixSze, SzeX * PixSze, I * PixSze
   Next I
   For I = 0 To SzeX - 1
      APIline picEd.hdc, 0, 0, 8, I * PixSze, 0, I * PixSze, SzeY * PixSze
   Next I
End Sub

' draw rounded rectangle
Private Sub DrawRRect(ByVal hdc As Integer, ByVal Color As Integer)
   Dim hPn As Long, hPnOld As Long
   Dim hBr As Long, hBrOld As Long
   
   If Action = A_FRRECT Then
      hBr = CreateSolidBrush(Color Xor &H1000000) ' Xor &H1000000 --> color = pallet-index
      hBrOld = SelectObject(hdc, hBr)
      Else
      hPn = CreatePen(0, PenWidth, Color Xor &H1000000)
      hPnOld = SelectObject(hdc, hPn)
      End If
      
   RoundRect hdc, X1, Y1, X2, Y2, _
            (X2 - X1) * Rounding / 100, _
            (Y2 - Y1) * Rounding / 100
            
   If Action = A_FRRECT Then
      SelectObject hdc, hBrOld
      DeleteObject hBr
      Else
      SelectObject hdc, hPnOld
      DeleteObject hPn
      End If
End Sub

' called in DoActionShade
Private Sub DrawShade(Color As Long)
   picPat.Width = SzeX * 15
   picPat.Height = SzeY * 15
   APIrect picPat.hdc, 0, Color, Color, 0, 0, SzeX, SzeY
   BitBlt picEx.hdc, 0, 0, SzeX, SzeY, picPat.hdc, 0, 0, SRCINVERT
   BitBlt picEx.hdc, 0, 0, SzeX, SzeY, picMask.hdc, 0, 0, SRCAND
   BitBlt picEx.hdc, 0, 0, SzeX, SzeY, picPat.hdc, 0, 0, SRCINVERT
   picEx.Refresh
End Sub

' called in DoAction
Private Sub FillWithPattern(pic As Control, ByVal vX1 As Integer, ByVal vY1 As Integer, ByVal W As Integer, ByVal H As Integer, PicP As Control, ByVal W2 As Integer, ByVal H2 As Integer)
   Dim Wa As Integer, Ha As Integer
   Dim X As Integer, Y As Integer
   
   Wa = W \ W2 + 1
   Ha = H \ H2 + 1
   For Y = 0 To Ha
      For X = 0 To Wa
         BitBlt pic.hdc, vX1 + X * W2, vY1 + Y * H2, W2, H2, PicP.hdc, 0, 0, SRCCOPY
      Next X
   Next Y
   pic.Refresh
End Sub

' main drawing routine
Private Sub DoAction()
   Dim PD As Long ' drawwidth (dutch:Pen Dikte)
   
   Me.MousePointer = vbHourglass
   
   ' mask
   DrawMask
   
   ' shade exept for fill action
   If Shade > 0 And Action <> A_FILL Then DoActionShade
   
   ' size of rectangle = size of image when action is pen or fill
   If Action = A_PEN Or Action = A_FILL Then
      X1 = 0: X2 = SzeX: dX = X2
      Y1 = 0: Y2 = SzeY: dY = Y2
      End If
      
   ' fillings
   If FillType = VT_CLIPBRAS Then ' grid of tiled clipb. content
      If Clipboard.GetFormat(2) = False Then
         FillType = VT_FLOODRAS ' no image in clipboard
         Else
         On Error Resume Next
         picBuff.AutoSize = True
         picBuff.Picture = Clipboard.GetData()
         picBuff.AutoSize = False
         If Err <> 0 Then FillType = VT_FLOODRAS: Beep
         On Error GoTo 0
         End If
      End If
   Select Case FillType
      Case VT_NONE, VT_SOLID
         picPat.Width = SzeX * 15
         picPat.Height = SzeY * 15
         APIrect picPat.hdc, 0, CurCol, CurCol, 0, 0, SzeX, SzeY
      Case VT_FLOOD
         picPat.Width = (dX + PenWidth) * 15
         picPat.Height = (dY + PenWidth) * 15
         Flood8b picPat, ColorL, ColorR, FloodType
      Case VT_FLOODRAS
         picPat.Width = PSzeX * 15
         picPat.Height = PSzeY * 15
         Flood8b picPat, ColorL, ColorR, FloodType
         FillWithPattern picPat, 0, 0, SzeX, SzeY, picPat, PSzeX, PSzeY
      Case VT_CLIPBRAS
         picPat.Width = SzeX * 15
         picPat.Height = SzeY * 15
         BitBlt picPat.hdc, 0, 0, picBuff.ScaleWidth, picBuff.ScaleHeight, picBuff.hdc, 0, 0, SRCCOPY
         FillWithPattern picPat, 0, 0, SzeX, SzeY, picPat, picBuff.ScaleWidth, picBuff.ScaleHeight
   End Select
   
   ' fuse mask & fillings
   PD = PenWidth
   If Action = A_PEN And Action = A_FILL Then
      ' all
      BitBlt picEx.hdc, 0, 0, picEx.ScaleWidth, picEx.ScaleHeight, picPat.hdc, 0, 0, SRCINVERT
      Else
      ' part - selection
      BitBlt picEx.hdc, X1 - PD \ 2, Y1 - PD \ 2, dX + PD, dY + PD, picPat.hdc, 0, 0, SRCINVERT
      End If
   BitBlt picEx.hdc, 0, 0, picEx.ScaleWidth, picEx.ScaleHeight, picMask.hdc, 0, 0, SRCAND
   If Action = A_PEN And Action = A_FILL Then
      BitBlt picEx.hdc, 0, 0, picEx.ScaleWidth, picEx.ScaleHeight, picPat.hdc, 0, 0, SRCINVERT
      Else
      BitBlt picEx.hdc, X1 - PD \ 2, Y1 - PD \ 2, dX + PD, dY + PD, picPat.hdc, 0, 0, SRCINVERT
      End If
   picEx.Refresh
   
   Me.MousePointer = vbDefault
End Sub

' used in DoActionShade
Private Sub RotateAll(ByVal Index, ByVal Distance)
   Dim I As Long, hV As Long, hB As Long
   
   hV = picEx.hdc: hB = picBuff.hdc
   For I = 0 To Distance - 1
      Select Case Index
      Case 0:  BitBlt hB, 0, 0, 1, SzeY, hV, 0, 0, SRCCOPY
               BitBlt hV, 0, 0, SzeX - 1, SzeY, hV, 1, 0, SRCCOPY
               BitBlt hV, SzeX - 1, 0, 1, SzeY, hB, 0, 0, SRCCOPY
      
      Case 1:  BitBlt hB, 0, 0, 1, SzeY, hV, SzeX - 1, 0, SRCCOPY
               BitBlt hV, 1, 0, SzeX - 1, SzeY, hV, 0, 0, SRCCOPY
               BitBlt hV, 0, 0, 1, SzeY, hB, 0, 0, SRCCOPY
      
      Case 2:  BitBlt hB, 0, 0, SzeX, 1, hV, 0, 0, SRCCOPY
               BitBlt hV, 0, 0, SzeX, SzeY - 1, hV, 0, 1, SRCCOPY
               BitBlt hV, 0, SzeY - 1, SzeX, 1, hB, 0, 0, SRCCOPY
   
      Case 3:  BitBlt hB, 0, 0, SzeX, 1, hV, 0, SzeY - 1, SRCCOPY
               BitBlt hV, 0, 1, SzeX, SzeY - 1, hV, 0, 0, SRCCOPY
               BitBlt hV, 0, 0, SzeX, 1, hB, 0, 0, SRCCOPY
      End Select
   Next I
End Sub

' set new image size
Private Sub SetSize()
   Me.MousePointer = vbHourglass
   If prSzeX <> SzeX Or prSzeY <> SzeY Then ' different size?
      kadEx.Width = (SzeX + 10) * 15
      kadEx.Height = (SzeY + 10) * 15
      picEx.Width = SzeX * 15
      picEx.Height = SzeY * 15
      picMask.Width = picEx.Width
      picMask.Height = picEx.Height
      picBuff.Width = picEx.Width
      picBuff.Height = picEx.Height
      picUndo.Width = picEx.Width
      picUndo.Height = picEx.Height
      End If
   
   picEd.Width = (SzeX * PixSze) * 15
   picEd.Height = (SzeY * PixSze) * 15
   VSEd.SmallChange = PixSze * 15 ' adapt scrollbars
   HSEd.SmallChange = PixSze * 15
   VSEd.LargeChange = PixSze * 15 * 4
   HSEd.LargeChange = PixSze * 15 * 4
   
   Form_Resize
   DoEvents
   
   picEd.Scale (0, 0)-(SzeX, SzeY)
   'DoEvents
   Me.MousePointer = vbDefault
End Sub

' get content of picEx (real size) and stretch (scale) it into picEd
Private Sub SmallToBig()
   StretchBlt picEd.hdc, 0, 0, SzeX * PixSze, SzeY * PixSze, picEx.hdc, 0, 0, SzeX, SzeY, SRCCOPY
   If ShowGrid Then DrawGrid
   picEd.Refresh
End Sub

' insert image with Paste or Open
Private Sub ShowNewPic()
   Dim nX1 As Integer, nY1 As Integer, ndX As Integer, ndY As Integer

   frmClpMd.Show 1
   If OK = False Then Exit Sub
   BeforeUndo
   
   If AdaptTo = 0 And ClipMode = 0 Then ' all/cut
      nX1 = 0: nY1 = 0: ndX = SzeX: ndY = SzeY
      StretchBlt picEx.hdc, nX1, nY1, ndX, ndY, picBuff.hdc, 0, 0, ndX, ndY, SRCCOPY
      End If
   If AdaptTo = 0 And ClipMode = 1 Then ' all/stretch
      nX1 = 0: nY1 = 0: ndX = SzeX: ndY = SzeY
      StretchBlt picEx.hdc, nX1, nY1, ndX, ndY, picBuff.hdc, 0, 0, picBuff.ScaleWidth, picBuff.ScaleHeight, SRCCOPY
      End If
   If AdaptTo = 0 And ClipMode = 2 Then ' all/new size
      nX1 = 0: nY1 = 0: ndX = picBuff.Width / 15: ndY = picBuff.Height / 15
      If ndX > MaxSize Then ndX = MaxSize
      If ndY > MaxSize Then ndY = MaxSize
      SzeX = ndX: SzeY = ndY: Call SetSize
      StretchBlt picEx.hdc, nX1, nY1, ndX, ndY, picBuff.hdc, 0, 0, ndX, ndY, SRCCOPY
      End If
   If AdaptTo = 1 And ClipMode = 0 Then ' sel/cut
      nX1 = X1: nY1 = Y1: ndX = dX: ndY = dY
      StretchBlt picEx.hdc, nX1, nY1, ndX, ndY, picBuff.hdc, 0, 0, ndX, ndY, SRCCOPY
      End If
   If AdaptTo = 1 And ClipMode = 1 Then ' sel/stretch
      nX1 = X1: nY1 = Y1: ndX = dX: ndY = dY
      StretchBlt picEx.hdc, nX1, nY1, ndX, ndY, picBuff.hdc, 0, 0, picBuff.ScaleWidth, picBuff.ScaleHeight, SRCCOPY
      End If
   If AdaptTo = 1 And ClipMode = 2 Then ' sel/new size
      nX1 = X1: nY1 = Y1: ndX = picBuff.Width / 15: ndY = picBuff.Height / 15
      StretchBlt picEx.hdc, nX1, nY1, ndX, ndY, picBuff.hdc, 0, 0, ndX, ndY, SRCCOPY
      X1 = nX1: Y1 = nY1: dX = ndX: dY = ndY
      X2 = X1 + dX: Y2 = Y1 + dY
      shpBox.Left = X1: shpBox.Width = dX
      shpBox.Top = Y1: shpBox.Height = dY
      End If
   
   picEx.Refresh
   
   SmallToBig
End Sub

' most of the pictureboxes need the same chosen pallet
Private Sub SetPalToPics()
   Dim Ret As Long
   Dim I As Long ' counter color ID
   Dim X As Long, Y As Long
   Dim hPal As Long
   
   Clipboard.Clear
   Ret = OpenClipboard(hwnd)
   If Ret = 0 Then MsgBox "clipboard error": End
   hPal = CreatePalette(CustPal)
   Ret = SetClipboardData(CF_PALETTE, hPal)
   Ret = CloseClipboard()
   picPal.Picture = Clipboard.GetData(CF_PALETTE)
   picEx.Picture = Clipboard.GetData(CF_PALETTE): APIcls picEx, BGCol
   picUndo.Picture = Clipboard.GetData(CF_PALETTE): APIcls picUndo, BGCol
   picEd.Picture = Clipboard.GetData(CF_PALETTE): APIcls picEd, BGCol
   picMse.Picture = Clipboard.GetData(CF_PALETTE): APIcls picMse, BGCol
   picPat.Picture = Clipboard.GetData(CF_PALETTE)
   Ret = DeleteObject(hPal)
   For I = 0 To 255
      X = 1 + (I Mod 8) * 8
      Y = 1 + (I \ 8) * 8
      APIrect picPal.hdc, 0, I, I, X, Y, X + 6, Y + 6
   Next I

End Sub

Private Sub cmdCopy_Click()
   Dim bW As Integer, bH As Integer
   
   Clipboard.Clear
   Clipboard.SetData picEx.Picture
   If Action = A_SELECT Then
      ' cut out the piece & copy into picBuff
      picBuff.Picture = LoadPicture()
      picBuff.Picture = Clipboard.GetData(9)
      bW = dX
      bH = dY
      picBuff.Width = bW * 15
      picBuff.Height = bH * 15
      StretchBlt picBuff.hdc, 0, 0, bW, bH, picEx.hdc, X1, Y1, bW, bH, SRCCOPY
      picBuff.Refresh
      Clipboard.SetData picBuff.Image
      Else
      ' copy All
      Clipboard.SetData picEx.Image
      End If
   picEd.SetFocus
End Sub

Private Sub cmdExit_Click()
   If MsgBox("Are you sure ?", vbOKCancel, "Exit PAINT256") = vbOK Then End
End Sub

Private Sub cmdHelp_Click()
   On Error Resume Next
   AppActivate "Help.wri", False
   If Err = 0 Then Exit Sub
   Err = 0
   Shell "write.exe " & App.Path & "\Help.wri", vbNormalFocus
   If Err <> 0 Then MsgBox Err.Description
   On Error GoTo 0
End Sub

Private Sub cmdInfo_Click()
   frmInfo.Show 1
End Sub

Private Sub cmdSetFillStyle_click()
   frmFill.Left = frmPnt.Left + cmdSetFillStyle.Left + cmdSetFillStyle.Width
   frmFill.Top = frmPnt.Top + cmdSetFillStyle.Top + cmdSetFillStyle.Height - frmFill.Height + 360
   frmFill.Show 1
   picEd.SetFocus
End Sub

' move (or rotate) all or selection
Private Sub cmdMve_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Static busy
   Dim hV As Long, hB As Long
   Dim Pz As Long
   
   If busy = True Then Exit Sub
   busy = True
   BeforeUndo
   
   ButtonDown = True: Pz = 750
   hV = picEx.hdc
   hB = picBuff.hdc
   Do While ButtonDown = True
       If Action = A_SELECT Then
      
          Select Case Index
             Case 0: BitBlt hB, 0, 0, 1, Y2 - Y1, hV, X1, Y1, SRCCOPY
                     BitBlt hV, X1, Y1, X2 - X1, Y2 - Y1, hV, X1 + 1, Y1, SRCCOPY
                     BitBlt hV, X2 - 1, Y1, 1, Y2 - Y1, hB, 0, 0, SRCCOPY
             Case 1: BitBlt hB, 0, 0, 1, Y2 - Y1, hV, X2 - 1, Y1, SRCCOPY
                     BitBlt hV, X1 + 1, Y1, X2 - X1 - 1, Y2 - Y1, hV, X1, Y1, SRCCOPY
                     BitBlt hV, X1, Y1, 1, Y2 - Y1, hB, 0, 0, SRCCOPY
             Case 2: BitBlt hB, 0, 0, X2 - X1, 1, hV, X1, Y1, SRCCOPY
                     BitBlt hV, X1, Y1, X2 - X1, Y2 - Y1, hV, X1, Y1 + 1, SRCCOPY
                     BitBlt hV, X1, Y2 - 1, X2 - X1, 1, hB, 0, 0, SRCCOPY
             Case 3: BitBlt hB, 0, 0, X2 - X1, 1, hV, X1, Y2 - 1, SRCCOPY
                     BitBlt hV, X1, Y1 + 1, X2 - X1, Y2 - Y1 - 1, hV, X1, Y1, SRCCOPY
                     BitBlt hV, X1, Y1, X2 - X1, 1, hB, 0, 0, SRCCOPY
          End Select
         
          Else
          RotateAll Index, 1
          End If
      
       picEx.Refresh
       SmallToBig
       
       Pause Pz: If Pz > 200 Then Pz = Pz * 0.8
   Loop
   busy = False
End Sub

Private Sub cmdMve_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   ButtonDown = False
   kadMove.Visible = False
   picEd.SetFocus
End Sub

Private Sub cmdOpen_click()
   Dim file As String
   
   file = file & "BitMaP File (*.bmp)|*.BMP"
   file = file & "|Icon File (*.ico)|*.ICO"
   file = file & "|Alle Files (*.*)|*.*"
   CD_File.CancelError = True
   CD_File.Filter = file
   CD_File.FilterIndex = 0
   CD_File.FileName = ""
   CD_File.DialogTitle = "Open File"
   CD_File.Flags = OFN_FILEMUSTEXIST
   
   On Error Resume Next
   CD_File.ShowOpen
   If Err <> 0 Then Exit Sub
   
   file = CD_File.FileName
   picBuff.AutoSize = True
   picBuff.Picture = LoadPicture(file)
   picBuff.AutoSize = False
   Action = A_SELECT
   pshAct(A_SELECT).Value = True
   ShowNewPic
   picEd.SetFocus
End Sub

Private Sub cmdPallet_Click()
   frmPal.Show 1
   If OK = False Then Exit Sub
   Me.MousePointer = vbHourglass
   DrawMouse
   SmallToBig
   Me.MousePointer = vbDefault
   picEd.SetFocus
End Sub

Private Sub cmdPalShift_Click()
   Static offs
   Dim ipb As String
   Dim ID As Integer, off As Integer
   Dim X As Integer, Y As Integer
   Dim StBdX As Long, StBFC As Long
   
   ipb = InputBox("Amount to add to the color-index ?", "Shift color-index", "16")
   If ipb = "" Then Exit Sub
   
   Me.MousePointer = vbHourglass
   StBdX = StatusBar.ScaleWidth
   StBFC = StatusBar.ForeColor
   
   offs = Val(ipb)
   If offs = 0 Then off = 8
   BeforeUndo
   If Action = A_SELECT Then
      For Y = Y1 To Y2 - 1
        For X = X1 To X2 - 1
          ID = IsColorID(picEx.Point(X, Y), 0)
          If ID > 15 Then
             ID = (ID + offs) Mod 256
             If ID < 16 Then ID = ID + 16
             picEx.PSet (X, Y), ColorSet(ID)
             End If
        Next X
        StatusBar.Line (0, 4)-((Y - Y1) / dX * StBdX, 16), StBFC, BF
      Next Y
      Else
      For Y = 0 To SzeY - 1
        For X = 0 To SzeX - 1
          ID = IsColorID(picEx.Point(X, Y), 0)
          If ID > 15 Then
             ID = (ID + offs) Mod 256
             If ID < 16 Then ID = ID + 16
             picEx.PSet (X, Y), ColorSet(ID)
             End If
        Next X
        StatusBar.Line (0, 4)-(Y / SzeY * StBdX, 16), StBFC, BF
      Next Y
      
      End If
   
   StatusBar.Line (0, 4)-(dY, 16), StBFC, BF
   SmallToBig
   StatusBar.Cls
   Me.MousePointer = vbDefault
   picEd.SetFocus
End Sub

Private Sub cmdPaste_click()
   If Clipboard.GetFormat(2) = False Then MsgBox "No image present in clipboard!": Exit Sub
   picBuff.AutoSize = True
   picBuff.Picture = Clipboard.GetData()
   picBuff.AutoSize = False
   DoEvents
   ShowNewPic
   Action = A_SELECT
   pshAct(Action).Value = True
   picEd.SetFocus
End Sub

Private Sub cmdDrawWidth_Click()
   kadWidths.Visible = True
End Sub

Private Sub cmdPictures_Click()
   frmBMP.Show
   If frmBMP.WindowState = 1 Then frmBMP.WindowState = 0
End Sub

Private Sub cmdSave_Click()
   Dim Filter As String, file As String
   Dim bX1 As Integer, bY1 As Integer, bW As Integer, bH As Integer
   
   Filter = Filter & "BitMaP File (*.bmp)|*.BMP"
   Filter = Filter & "|Alle Files (*.*)|*.*"
   CD_File.CancelError = True
   CD_File.Filter = Filter
   CD_File.FilterIndex = 0
   CD_File.FileName = "*.BMP"
   CD_File.DialogTitle = "Save File"
   CD_File.Flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
   CD_File.DefaultExt = ".BMP"
   
   On Error Resume Next
   CD_File.ShowSave
   If Err <> 0 Then Exit Sub
   
   file = CD_File.FileName
   Me.MousePointer = vbHourglass
   Clipboard.Clear
   Clipboard.SetData picEx.Picture
   picBuff.Picture = LoadPicture()
   picBuff.Picture = Clipboard.GetData(9)
   If Action = A_SELECT Then
      bX1 = X1: bY1 = Y1: bW = dX: bH = dY
      Else
      bX1 = 0: bY1 = 0: bW = SzeX: bH = SzeY
      End If
   picBuff.Width = bW * 15
   picBuff.Height = bH * 15
   StretchBlt picBuff.hdc, 0, 0, bW, bH, picEx.hdc, bX1, bY1, bW, bH, SRCCOPY
   picBuff.Refresh
   SavePicture picBuff.Image, file
   Me.MousePointer = vbDefault
   picEd.SetFocus
End Sub

Private Sub cmdSchades_click()
   frmSchad.Left = frmPnt.Left + cmdSchades.Left + cmdSchades.Width
   frmSchad.Top = frmPnt.Top + cmdSchades.Top + cmdSchades.Height - frmSchad.Height + 360
   frmSchad.Show 1
   picEd.SetFocus
End Sub

Private Sub cmdSize_click()
   BeforeUndo
   frmSize.Show 1
   If OK = True Then
      SetSize
      SmallToBig
      End If
   picEd.SetFocus
End Sub

Private Sub cmdMirrH_Click()
   Dim X As Integer
   
   Me.MousePointer = vbHourglass
   BeforeUndo
   picBuff.Cls
   If Action <> A_SELECT Then
      For X = 0 To SzeX
         BitBlt picBuff.hdc, SzeX - X - 1, 0, 1, SzeY, picEx.hdc, X, 0, SRCCOPY
      Next X
      picBuff.Refresh
      BitBlt picEx.hdc, 0, 0, SzeX, SzeY, picBuff.hdc, 0, 0, SRCCOPY
      Else
      For X = 0 To X2 - X1 - 1
         BitBlt picBuff.hdc, X, 0, 1, Y2 - Y1, picEx.hdc, X2 - X - 1, Y1, SRCCOPY
      Next X
      picBuff.Refresh
      BitBlt picEx.hdc, X1, Y1, X2 - X1, Y2 - Y1, picBuff.hdc, 0, 0, SRCCOPY
      End If
   
   SmallToBig
   Me.MousePointer = vbDefault
   picEd.SetFocus
End Sub

Private Sub cmdMirrV_Click()
   Dim Y As Integer

   Me.MousePointer = vbHourglass
   BeforeUndo
   picBuff.Cls
   If Action <> A_SELECT Then ' all or selection
      For Y = 0 To SzeY
         BitBlt picBuff.hdc, 0, SzeY - Y - 1, SzeX, 1, picEx.hdc, 0, Y, SRCCOPY
      Next Y
      picBuff.Refresh
      BitBlt picEx.hdc, 0, 0, SzeX, SzeY, picBuff.hdc, 0, 0, SRCCOPY
      Else
      For Y = 0 To Y2 - Y1 - 1
         BitBlt picBuff.hdc, 0, Y, X2 - X1, 1, picEx.hdc, X1, Y2 - Y - 1, SRCCOPY
      Next Y
      picBuff.Refresh
      BitBlt picEx.hdc, X1, Y1, X2 - X1, Y2 - Y1, picBuff.hdc, 0, 0, SRCCOPY
      End If
   
   SmallToBig
   Me.MousePointer = vbDefault
   picEd.SetFocus
End Sub

Private Sub cmdUndo_click()
   Dim Fl As Boolean
   
   Me.MousePointer = vbHourglass
   If prSzeX <> SzeX Then SzeX = prSzeX: Fl = True
   If prSzeY <> SzeY Then SzeY = prSzeY: Fl = True
   If Fl = True Then Call SetSize
   StretchBlt picEx.hdc, 0, 0, SzeX, SzeY, picUndo.hdc, 0, 0, SzeX, SzeY, SRCCOPY
   picEx.Refresh
   SmallToBig
   Me.MousePointer = vbDefault
   picEd.SetFocus
End Sub

Private Sub cmdVB_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   kadEx.Visible = True
End Sub

Private Sub cmdVB_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   kadEx.Visible = False
   picEd.SetFocus
End Sub

Private Sub cmdMove_Click()
   kadMove.Visible = True
End Sub

Private Sub cmdClear_click()
   Me.MousePointer = vbHourglass
   If Action = A_SELECT Then
      APIrect picEx.hdc, 0, BGCol, BGCol, X1, Y1, X2, Y2
      Else
      APIrect picEx.hdc, 0, BGCol, BGCol, 0, 0, SzeX, SzeY
      End If
   SmallToBig
   Me.MousePointer = vbDefault
   picEd.SetFocus
End Sub

Private Sub Form_Activate()
   If Loading = True Then
      SetSize
      SmallToBig
      Screen.MousePointer = vbDefault
      Loading = False
      End If
   Timer1.Enabled = True
End Sub

Sub Form_Deactivate()
   Timer1.Enabled = False
End Sub

Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   Dim mFl As Boolean
   
   Select Case KeyCode
      Case Asc("C"): KeyCode = 0: cmdCopy_Click
      Case Asc("E"): KeyCode = 0: cmdExit_Click
      Case Asc("O"): KeyCode = 0: cmdOpen_click
      Case Asc("P"): KeyCode = 0: cmdPaste_click
      Case Asc("B"): KeyCode = 0: cmdSave_Click
      Case Asc("A"): KeyCode = 0: cmdSize_click
      Case Asc("U"): KeyCode = 0: cmdUndo_click
      Case Asc("W"): KeyCode = 0: cmdClear_click
      Case Asc("S"): KeyCode = 0: cmdSchades_click
      Case Asc("V"): KeyCode = 0: cmdSetFillStyle_click
      
      Case Asc(" "): KeyCode = 0
         Select Case Action
         Case A_BEVEL: DoBevel picEx, BevelType: SmallToBig
         Case A_FILL
         Case Else: DoAction: SmallToBig
         End Select
      
      Case vbKeyLeft: mFl = True
         If Shift = 1 Then X1 = X1 + 1 Else X1 = X1 - 1
         If Shift = 2 Then X2 = X2 - 1
      Case vbKeyRight: mFl = True
         If Shift = 1 Then X2 = X2 - 1 Else X2 = X2 + 1
         If Shift = 2 Then X1 = X1 + 1
      Case vbKeyUp: mFl = True
         If Shift = 1 Then Y1 = Y1 + 1 Else Y1 = Y1 - 1
         If Shift = 2 Then Y2 = Y2 - 1
      Case vbKeyDown: mFl = True
         If Shift = 1 Then Y2 = Y2 - 1 Else Y2 = Y2 + 1
         If Shift = 2 Then Y1 = Y1 + 1
   End Select
   
   If mFl = True Then
      KeyCode = 0
      shpBox.Left = X1: shpBox.Top = Y1
      dX = Abs(X2 - X1): dY = Abs(Y2 - Y1)
      shpBox.Width = dX: shpBox.Height = dY
      shpBox.Visible = True
      End If
   
   If Shift <> 0 And KeyCode = Asc("M") Then ' hidden feature for testing purposes
      If picMask.Visible = True Then
         picMask.Visible = False
         picPat.Visible = False
         Else
         picMask.Visible = True
         picPat.Visible = True
         End If
      End If
   If Shift <> 0 And KeyCode = Asc("K") Then
      kadEx.Visible = True
      End If
End Sub

Private Sub Form_Load()
   Loading = True
   Screen.MousePointer = vbHourglass
   
   Me.Caption = "Paint By: Enrique A. Flores B. and Kew Lung!"
   
   kadO.ZOrder: kadEx.ZOrder: kadMove.ZOrder
   kadSh.ZOrder: kadWidths.ZOrder: kadRHCK.ZOrder
   
   BaseColorStart
   BGCol = 7
   picBG.BackColor = ColorSet(BGCol)
   ColorL = 0
   ColorR = BGCol
   CurCol = 0
   SetPalToPics
   DrawMouse
   shpMseL.Left = 0: shpMseL.Top = 0
   shpMseR.Left = 7 * 8: shpMseR.Top = 0
   
   PixSze = 8
   SzeX = 128: SzeY = 64: prSzeX = 127: prSzeY = 63
   PSzeX = 8: PSzeY = 8
   ShowGrid = True
   PenWidth = 1
   
   X1 = 8: X2 = 24: Y1 = 8: Y2 = 24: dX = 16: dY = 16
   lblX = 0: lblY = 0
   lblX1 = X1 / PixSze: lblY1 = Y1 / PixSze
   lblX2 = X2 / PixSze: lblY2 = Y2 / PixSze
   lblDX = dX / PixSze: lblDY = dY / PixSze
   shpBox.Left = X1: shpBox.Top = Y1
   shpBox.Width = dX: shpBox.Height = dY
   
   ShadeDis = 1
   
   FormulaPts = 4
   FormulaGplus = 120
   FormulaAngle = 270
   
   FillType = VT_SOLID
   
   Text = "PAINT 256"
   TextFntNm = "Times New Roman": TextSize = 9
   TextdX = 58: TextdY = 15
   
   ArrowAngle = 0
   ArrowWidth = 2
   ArrowHeadW = 5
   ArrowFilled = True
   
   Action = A_PEN
End Sub

Private Sub Form_Resize()
   If WindowState = 1 Then Exit Sub
   If Width < 7080 Then Width = 7080
   If Height < 6210 Then Height = 6210
   StatusBar.Top = ScaleHeight - StatusBar.Height - 45
   StatusBar.Width = ScaleWidth - StatusBar.Left * 3
   kadO.Width = ScaleWidth - kadO.Left - 60
   kadO.Height = StatusBar.Top - kadO.Top - 60
   CheckScrolls picEd, kadEd, kadO, HSEd, VSEd
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Private Sub HSEd_Change()
   picEd.Left = -HSEd.Value
End Sub

Private Sub HSEd_Scroll()
   picEd.Left = -HSEd.Value
End Sub

Private Sub kadWidths_Paint()
   BevelObject kadWidths, 15, 15, kadWidths.ScaleWidth - 15, kadWidths.ScaleHeight - 15, 0
End Sub

Private Sub kadRHCK_Paint()
   BevelObject kadRHCK, 15, 15, kadRHCK.ScaleWidth - 15, kadRHCK.ScaleHeight - 15, 0
End Sub

Private Sub kadSh_Paint()
   BevelObject kadSh, 15, 15, kadSh.ScaleWidth - 15, kadSh.ScaleHeight - 15, 0
End Sub

Private Sub kadEx_Paint()
   BevelObject kadEx, 15, 15, kadEx.ScaleWidth - 15, kadEx.ScaleHeight - 15, 0
End Sub

Private Sub kadMove_Paint()
   BevelObject kadMove, 15, 15, kadMove.ScaleWidth - 15, kadMove.ScaleHeight - 15, 0
End Sub

Private Sub picRounding_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Rounding = 100 / picRounding.ScaleWidth * X
   picRounding.Cls
   picRounding.Line (0, 0)-(X, 195), QBColor(9), BF
   picRHCK.Refresh
   APIrrect picRHCK.hdc, 0, 0, BGCol, 0, 0, picRHCK.ScaleWidth, picRHCK.ScaleHeight, Rounding
End Sub

Private Sub picRounding_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = 1 Then
      Rounding = 100 / picRounding.ScaleWidth * X
      If Rounding > 100 Then Rounding = 100
      If Rounding < 0 Then Rounding = 0
      picRounding.Cls
      picRounding.Line (0, 0)-(X, 195), QBColor(9), BF
      picRHCK.Refresh
      APIrrect picRHCK.hdc, 0, 0, BGCol, 0, 0, picRHCK.ScaleWidth, picRHCK.ScaleHeight, Rounding
      End If
End Sub

Private Sub picRounding_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   kadRHCK.Visible = False
End Sub

Private Sub picRounding_Paint()
   picRounding.Cls
   picRounding.Line (0, 0)-(Rounding * picRounding.ScaleWidth / 100, 195), QBColor(9), BF
End Sub

' hide all 'popup frames'
Private Sub picEd_GotFocus()
   kadWidths.Visible = False
   kadRHCK.Visible = False
   kadSh.Visible = False
   kadMove.Visible = False
End Sub

Private Sub picEd_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim ID As Long
   
   mX = Int(X): mY = Int(Y)
   Select Case Shift
      Case 0, 1, 2 ' all except Alt key
         If Button = 1 Then
            ' if clicking same color as in mouse -> reverse to BG
            If picEx.Point(mX, mY) = ColorSet(ColorL) Then CurCol = BGCol Else CurCol = ColorL
            Else
            CurCol = ColorR
            End If
         BeforeUndo
         If Action = A_SELECT Then
            If X1 < mX And mX < X2 And Y1 < mY And mY < Y2 Then
               ' inside current box
               mX1 = mX: mY1 = mY: mX2 = mX + 1: mY2 = mY + 1
               picBuff.Width = dX * 15
               picBuff.Height = dY * 15
               ' copy content box to buffer
               StretchBlt picBuff.hdc, 0, 0, dX, dY, picEx.hdc, X1, Y1, dX, dY, SRCCOPY
               picBuff.Refresh
               If Shift = 2 Then ' moving so clear content
                  picEx.Line (X1, Y1)-(X2 - 1, Y2 - 1), ColorSet(BGCol), BF
                  SmallToBig
                  End If
               MovingBox = True
               picEd.MousePointer = vbSizeAll
               Exit Sub
               End If
            End If
         ' outside current box --> new box
         X1 = mX: Y1 = mY: X2 = mX + 1: Y2 = mY + 1
         dX = 1: dY = 1
         NewBox = True ' all actions have a box, freehand also
         Select Case Action
            Case A_PEN
               picMask.Cls ' mask will be made on the fly (=exception)
               DrawPixel mX, mY, CurCol
            Case A_LINE
               Line1.X1 = mX: Line1.Y1 = mY
               Line1.X2 = mX + 1: Line1.Y2 = mY + 1
               Line1.Visible = True
            Case A_TEXT
               X1 = Int(mX - Abs(TextdX) / 2) ' center around mouse pos.
               Y1 = Int(mY - Abs(TextdY) / 2)
               X2 = Int(mX + Abs(TextdX) / 2)
               Y2 = Int(mY + Abs(TextdY) / 2)
               mX1 = mX: mY1 = mY: mX2 = mX: mY2 = mY
               dX = Abs(TextdX): dY = Abs(TextdY)
               shpBox.Left = X1: shpBox.Top = Y1
               shpBox.Width = dX: shpBox.Height = dY
               shpBox.Visible = True
               MovingBox = True
            Case Else
               shpBox.Left = X1: shpBox.Top = Y1
               shpBox.Width = dX: shpBox.Height = dY
               shpBox.Visible = True
         End Select
         
      Case 4 ' alt -> color under mouse pointer is (left/right) mouse color
         ID = IsColorID(picEx.Point(mX, mY), 0)
         X = 2 + (ID Mod 8) * 8
         Y = 2 + (ID \ 8) * 8
         picPal_MouseDown Button, 0, X, Y
   End Select
End Sub

Private Sub PicEd_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   mX = Int(X): mY = Int(Y)
   
   If Button = 0 Then Exit Sub ' not dragging
   
   AutoScroll kadEd, mX, mY, HSEd, VSEd ' mouse moves out of range, so adapt scrolls.
   
   If MovingBox = False And Action <> A_PEN Then
      ' adapt rubber box size
      If Abs(mX - X1) < Abs(mX - X2) Then X1 = mX: shpBox.Left = X1 Else X2 = mX
      If Abs(mY - Y1) < Abs(mY - Y2) Then Y1 = mY: shpBox.Top = Y1 Else Y2 = mY
      dX = Abs(X2 - X1): dY = Abs(Y2 - Y1)
      shpBox.Width = dX: shpBox.Height = dY
      End If
   
   Select Case Action
      Case A_SELECT
         If MovingBox = True Then
            mX2 = mX: mY2 = mY
            shpBox.Left = X1 + (mX2 - mX1)
            shpBox.Top = Y1 + (mY2 - mY1)
            End If
      Case A_PEN
         DrawPixel mX, mY, CurCol ' show a temporal drawing
      Case A_LINE
         Line1.X2 = mX: Line1.Y2 = mY
      Case A_TEXT
         mX2 = mX: mY2 = mY
         shpBox.Left = X1 + (mX2 - mX1)
         shpBox.Top = Y1 + (mY2 - mY1)
   End Select
   
End Sub

Private Sub picEd_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Shift = 4 Then Exit Sub ' Alt -> nothing to do here
   mX = Int(X): mY = Int(Y)
   If MovingBox = True Then
      ' calc new position
      X1 = X1 + (mX2 - mX1): Y1 = Y1 + (mY2 - mY1)
      X2 = X2 + (mX2 - mX1): Y2 = Y2 + (mY2 - mY1)
      mX2 = 0: mY2 = 0
      dX = Abs(X2 - X1): dY = Abs(Y2 - Y1)
      picEd.MousePointer = vbDefault
      If Shift <> 0 Then
         ' copy content of buffer into new position
         StretchBlt picEx.hdc, X1, Y1, dX, dY, picBuff.hdc, 0, 0, dX, dY, SRCCOPY
         picEx.Refresh
         SmallToBig
         End If
      End If
   
   Select Case Action
      Case A_SELECT ' no more actions to perform here
      Case A_TEXT
         X1 = shpBox.Left: Y1 = shpBox.Top
         dX = shpBox.Width: dY = shpBox.Height
         X2 = X1 + dX: Y2 = Y1 + dY
         Call DoAction
      Case A_BEVEL
         Call DoBevel(picEx, BevelType) ' inset/raised frames have their own routine
      Case Else
         Call DoAction
   End Select
   
   If Action > A_SELECT Then
      Call SmallToBig   ' show the result
      shpBox.Visible = False
      Line1.Visible = False
      End If
   MovingBox = False
   NewBox = False
   
End Sub

' replace left mouse color with right mouse color
Private Sub picMse_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim StBdX As Long, StBFC As Long
   
   Me.MousePointer = vbHourglass
   StBdX = StatusBar.ScaleWidth
   StBFC = StatusBar.ForeColor
   BeforeUndo
   If Action <> A_SELECT Then
      ' all
      For Y = 0 To SzeY - 1
         For X = 0 To SzeX - 1
            If picEx.Point(X, Y) = ColorSet(ColorL) Then picEx.PSet (X, Y), ColorSet(ColorR)
         Next X
         StatusBar.Line (0, 4)-(Y / SzeY * StBdX, 16), StBFC, BF
      Next Y
      Else
      ' selection
      For Y = Y1 To Y2 - 1
         For X = X1 To X2 - 1
            If picEx.Point(X, Y) = ColorSet(ColorL) Then picEx.PSet (X, Y), ColorSet(ColorR)
         Next X
         StatusBar.Line (0, 4)-((Y - Y1) / dX * StBdX, 16), StBFC, BF
      Next Y
      End If
   
   StatusBar.Line (0, 4)-(dY, 16), StBFC, BF
   
   SmallToBig
   StatusBar.Cls
   Me.MousePointer = vbDefault
   
End Sub

' choose left & right mouse colors
Private Sub picPal_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim ID As Integer
   ID = (Y \ 8) * 8 + X \ 8
   If Shift = 0 Then
      If Button = 1 Then
         ColorL = ID ' new left color
         lblPalIDL.Caption = ID
         shpMseL.Left = Fix(X \ 8) * 8
         shpMseL.Top = Fix(Y \ 8) * 8
         End If
      If Button = 2 Then
         ColorR = ID ' new right color
         lblPalIDR.Caption = ID
         shpMseR.Left = Fix(X \ 8) * 8
         shpMseR.Top = Fix(Y \ 8) * 8
         End If
      DrawMouse
      CurCol = ID
      Else
      BGCol = ID ' background color change
      picBG.BackColor = ColorSet(ID)
      End If
End Sub

Private Sub picRHCK_Paint()
   APIrrect picRHCK.hdc, 0, 0, BGCol, 0, 0, picRHCK.ScaleWidth, picRHCK.ScaleHeight, Rounding
End Sub

' select an instrument/Action
Private Sub pshAct_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Select Case Button
   Case 1
      Action = Index
      shpBox.Visible = False
      Select Case Action
         Case A_SELECT:
            shpBox.Shape = 0
            shpBox.Left = X1: shpBox.Top = Y1
            shpBox.Width = dX: shpBox.Height = dY
            shpBox.Visible = True
         Case Else: shpBox.Shape = 0
      End Select
   Case 2
      OK = False
      Select Case Index
         Case A_FORMULA: frmForm.Show 1
         Case A_RRECT: kadRHCK.Visible = True
         Case A_FRRECT: kadRHCK.Visible = True
         Case A_TEXT: frmText.Show 1, Me
         Case A_BEVEL: kadSh.Visible = True
         Case A_ARROW: frmArrows.Show 1, Me
      End Select
      If OK = True Then Action = Index: pshAct(Index).Value = True
   End Select
End Sub

Private Sub pshWidth_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   cmdDrawWidth.Picture = pshWidth(Index).Picture
   kadWidths.Visible = False
   PenWidth = Index + 1
   picMask.DrawWidth = PenWidth
End Sub

Private Sub pshSh_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   pshAct(A_BEVEL).Picture = pshSh(Index).Picture
   pshAct(A_BEVEL).Value = True
   kadSh.Visible = False
   BevelType = Index
   Action = A_BEVEL
End Sub

' update current coordinates
Private Sub Timer1_Timer()
   lblX = mX: lblY = mY
   lblX1 = X1: lblY1 = Y1
   lblX2 = X2: lblY2 = Y2
   lblDX = dX: lblDY = dY
End Sub

Private Sub VSEd_Change()
   picEd.Top = -VSEd.Value
End Sub

Private Sub VSEd_Scroll()
   picEd.Top = -VSEd.Value
End Sub

