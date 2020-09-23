VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DM Color Picker"
   ClientHeight    =   4185
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   9900
   Icon            =   "PalMaker.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   9900
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdinvert 
      Caption         =   "Invert Colour"
      Height          =   1185
      Left            =   8325
      Picture         =   "PalMaker.frx":0CCA
      TabIndex        =   61
      Top             =   2205
      Width           =   1455
   End
   Begin VB.CommandButton cmdrand2 
      Caption         =   "Rand"
      Height          =   270
      Index           =   2
      Left            =   4275
      TabIndex        =   60
      Top             =   3090
      Width           =   660
   End
   Begin VB.CommandButton cmdrand2 
      Caption         =   "Rand"
      Height          =   270
      Index           =   1
      Left            =   4275
      TabIndex        =   59
      Top             =   2670
      Width           =   660
   End
   Begin VB.CommandButton cmdrand2 
      Caption         =   "Rand"
      Height          =   270
      Index           =   0
      Left            =   4275
      TabIndex        =   58
      Top             =   2205
      Width           =   660
   End
   Begin VB.CommandButton cmdcopy 
      Caption         =   "Copy"
      Height          =   330
      Left            =   7515
      TabIndex        =   57
      Top             =   1455
      Width           =   2145
   End
   Begin VB.TextBox txtShCol 
      Height          =   300
      Left            =   7515
      Locked          =   -1  'True
      TabIndex        =   56
      Top             =   1050
      Width           =   2145
   End
   Begin VB.ListBox lstOp 
      Height          =   645
      Left            =   7515
      TabIndex        =   55
      Top             =   255
      Width           =   2145
   End
   Begin VB.CommandButton cmdScreen 
      Caption         =   "Colour from Screen"
      Height          =   1185
      Left            =   6705
      Picture         =   "PalMaker.frx":0FD4
      Style           =   1  'Graphical
      TabIndex        =   54
      Top             =   2205
      Width           =   1455
   End
   Begin VB.CommandButton cmdRand 
      Caption         =   "Random RGB Values"
      Height          =   1185
      Left            =   5085
      TabIndex        =   52
      Top             =   2205
      Width           =   1455
   End
   Begin VB.TextBox txtrgb 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   2
      Left            =   3675
      TabIndex        =   48
      Top             =   3068
      Width           =   480
   End
   Begin VB.TextBox txtrgb 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   1
      Left            =   3675
      TabIndex        =   47
      Top             =   2648
      Width           =   480
   End
   Begin VB.TextBox txtrgb 
      Alignment       =   2  'Center
      Height          =   315
      Index           =   0
      Left            =   3675
      TabIndex        =   46
      Top             =   2205
      Width           =   480
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1635
      Left            =   6015
      MouseIcon       =   "PalMaker.frx":12DE
      MousePointer    =   99  'Custom
      ScaleHeight     =   109
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   93
      TabIndex        =   42
      Top             =   195
      Width           =   1400
   End
   Begin Project1.Line3D Line3D3 
      Height          =   90
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   159
   End
   Begin Project1.Line3D Line3D1 
      Height          =   1650
      Left            =   2025
      TabIndex        =   38
      Top             =   180
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   2910
      LineStyle       =   2
   End
   Begin VB.HScrollBar hsbgreen 
      Height          =   300
      Left            =   750
      Max             =   255
      TabIndex        =   37
      Top             =   2655
      Width           =   2760
   End
   Begin VB.HScrollBar hsbred 
      Height          =   300
      Left            =   750
      Max             =   255
      TabIndex        =   36
      Top             =   2205
      Width           =   2760
   End
   Begin VB.HScrollBar hsbblue 
      Height          =   300
      Left            =   750
      Max             =   255
      TabIndex        =   35
      Top             =   3075
      Width           =   2760
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   6
      Left            =   4935
      MouseIcon       =   "PalMaker.frx":15E8
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   34
      Top             =   195
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   5
      Left            =   4455
      MouseIcon       =   "PalMaker.frx":18F2
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   33
      Top             =   195
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   14
      Left            =   4935
      MouseIcon       =   "PalMaker.frx":1BFC
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   32
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   22
      Left            =   4935
      MouseIcon       =   "PalMaker.frx":1F06
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   31
      Top             =   1005
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   30
      Left            =   4935
      MouseIcon       =   "PalMaker.frx":2210
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   30
      Top             =   1410
      Width           =   480
   End
   Begin Project1.Flat2 Flat21 
      Height          =   1725
      Left            =   90
      TabIndex        =   28
      Top             =   150
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   3043
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   31
      Left            =   5415
      MouseIcon       =   "PalMaker.frx":251A
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   27
      Top             =   1410
      Width           =   480
   End
   Begin VB.PictureBox colview 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1620
      Left            =   120
      ScaleHeight     =   1620
      ScaleWidth      =   1860
      TabIndex        =   26
      Top             =   195
      Width           =   1860
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preview Colour"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   29
         Top             =   660
         Width           =   1485
      End
   End
   Begin MSComDlg.CommonDialog CDialog 
      Left            =   360
      Top             =   4860
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   29
      Left            =   4455
      MouseIcon       =   "PalMaker.frx":2824
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   25
      Top             =   1410
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   28
      Left            =   3975
      MouseIcon       =   "PalMaker.frx":2B2E
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   24
      Top             =   1410
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   27
      Left            =   3495
      MouseIcon       =   "PalMaker.frx":2E38
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   23
      Top             =   1410
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   26
      Left            =   3015
      MouseIcon       =   "PalMaker.frx":3142
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   22
      Top             =   1410
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   25
      Left            =   2535
      MouseIcon       =   "PalMaker.frx":344C
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   21
      Top             =   1410
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   24
      Left            =   2055
      MouseIcon       =   "PalMaker.frx":3756
      MousePointer    =   99  'Custom
      ScaleHeight     =   405
      ScaleWidth      =   480
      TabIndex        =   20
      Top             =   1410
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   23
      Left            =   5415
      MouseIcon       =   "PalMaker.frx":3A60
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   19
      Top             =   1005
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   21
      Left            =   4455
      MouseIcon       =   "PalMaker.frx":3D6A
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   18
      Top             =   1005
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   20
      Left            =   3975
      MouseIcon       =   "PalMaker.frx":4074
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   17
      Top             =   1005
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   19
      Left            =   3495
      MouseIcon       =   "PalMaker.frx":437E
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   16
      Top             =   1005
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   18
      Left            =   3015
      MouseIcon       =   "PalMaker.frx":4688
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   15
      Top             =   1005
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   17
      Left            =   2535
      MouseIcon       =   "PalMaker.frx":4992
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   14
      Top             =   1005
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   16
      Left            =   2055
      MouseIcon       =   "PalMaker.frx":4C9C
      MousePointer    =   99  'Custom
      ScaleHeight     =   405
      ScaleWidth      =   480
      TabIndex        =   13
      Top             =   1005
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   15
      Left            =   5415
      MouseIcon       =   "PalMaker.frx":4FA6
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   12
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   13
      Left            =   4455
      MouseIcon       =   "PalMaker.frx":52B0
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   11
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   12
      Left            =   3975
      MouseIcon       =   "PalMaker.frx":55BA
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   10
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   11
      Left            =   3495
      MouseIcon       =   "PalMaker.frx":58C4
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   9
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   10
      Left            =   3015
      MouseIcon       =   "PalMaker.frx":5BCE
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   8
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   9
      Left            =   2535
      MouseIcon       =   "PalMaker.frx":5ED8
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   7
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   8
      Left            =   2055
      MouseIcon       =   "PalMaker.frx":61E2
      MousePointer    =   99  'Custom
      ScaleHeight     =   405
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   600
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   7
      Left            =   5415
      MouseIcon       =   "PalMaker.frx":64EC
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   195
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   4
      Left            =   3975
      MouseIcon       =   "PalMaker.frx":67F6
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   4
      Top             =   195
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   3
      Left            =   3495
      MouseIcon       =   "PalMaker.frx":6B00
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   3
      Top             =   195
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   2
      Left            =   3015
      MouseIcon       =   "PalMaker.frx":6E0A
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   2
      Top             =   195
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   1
      Left            =   2535
      MouseIcon       =   "PalMaker.frx":7114
      MousePointer    =   99  'Custom
      ScaleHeight     =   729
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   1
      Top             =   195
      Width           =   480
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   405
      Index           =   0
      Left            =   2055
      MouseIcon       =   "PalMaker.frx":741E
      MousePointer    =   99  'Custom
      ScaleHeight     =   405
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   195
      Width           =   480
   End
   Begin Project1.Line3D Line3D2 
      Height          =   1650
      Left            =   5940
      TabIndex        =   39
      Top             =   165
      Width           =   90
      _ExtentX        =   159
      _ExtentY        =   2910
      LineStyle       =   2
   End
   Begin Project1.Line3D Line3D4 
      Height          =   90
      Left            =   0
      TabIndex        =   41
      Top             =   2040
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   159
   End
   Begin Project1.Line3D Line3D5 
      Height          =   90
      Left            =   0
      TabIndex        =   53
      Top             =   3675
      Width           =   2925
      _ExtentX        =   5159
      _ExtentY        =   159
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "DM Color Picker Version 2.05"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   105
      TabIndex        =   62
      Top             =   3825
      Width           =   2100
   End
   Begin VB.Label lblblue 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   120
      TabIndex        =   51
      Top             =   5085
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblgreen 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   120
      TabIndex        =   50
      Top             =   4485
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblred 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Left            =   8280
      TabIndex        =   49
      Top             =   1965
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blue"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   120
      TabIndex        =   45
      Top             =   3120
      Width           =   360
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Green"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   120
      TabIndex        =   44
      Top             =   2700
      Width           =   510
   End
   Begin VB.Label lblrgb 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Red"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   0
      Left            =   120
      TabIndex        =   43
      Top             =   2250
      Width           =   315
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuopen 
         Caption         =   "&Open Pallet"
      End
      Begin VB.Menu mnunewpal 
         Caption         =   "&Create new Pallet"
      End
      Begin VB.Menu mnuabout 
         Caption         =   "&About"
      End
      Begin VB.Menu mnublank 
         Caption         =   "-"
      End
      Begin VB.Menu mnuexit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim isOn As Boolean, FirstTime As Boolean, StrRgb As String, nAction As Integer, LsIdx As Integer

Sub GetColour()
Dim iColor As Long
    Select Case LsIdx
        Case 0
            txtShCol.Text = DoHex(colview.BackColor, mCPlus)
        Case 1
            txtShCol.Text = DoHex(colview.BackColor, mDelphi)
        Case 2
            txtShCol.Text = DoHex(colview.BackColor, mVB)
        Case 3
            txtShCol.Text = colview.BackColor
        Case 4
            txtShCol.Text = txtrgb(0) & "," & txtrgb(1) & "," & txtrgb(2)
        Case 5
            iColor = txtrgb(2) * 256 * 256 + txtrgb(1) * 256 + txtrgb(0)
            txtShCol.Text = Dec2Web(iColor)
    End Select
    
End Sub

Function LoadPallet(lzFile As String) As Long
Dim tFile As Long
Dim sHead As String * 5

    tFile = FreeFile
    Open lzFile For Binary As #tFile
        Get #tFile, , sHead
        If Len(sHead) < 5 Then
            LoadPallet = 0
            Close #tFile
            Exit Function
        ElseIf Not Asc(Right(sHead, 1)) = 4 Then
            LoadPallet = 0
            Close #tFile
            Exit Function
        ElseIf FileLen(lzFile) < 137 Then
            LoadPallet = 0
            Close #tFile
            Exit Function
        Else
            Get #tFile, , Pallet
            LoadPallet = 1
        End If
    Close #tFile
    sHead = ""
    
End Function

Private Sub cmdcopy_Click()
    Clipboard.Clear
    Clipboard.SetText txtShCol.Text
End Sub

Private Sub cmdinvert_Click()
colview.BackColor = InvertColour(colview.BackColor)

    LongToRgb colview.BackColor
    txtrgb(0).Text = T_RGB.Red
    txtrgb(1).Text = T_RGB.Green
    txtrgb(2).Text = T_RGB.Blue
    Gradient Picture2, Val(txtrgb(0)), Val(txtrgb(1)), Val(txtrgb(2)), True

End Sub

Private Sub cmdrand_Click()
    Randomize
    txtrgb(0).Text = Int(Rnd * 255) + 1
    txtrgb(1).Text = Int(Rnd * 255) + 1
    txtrgb(2).Text = Int(Rnd * 255) + 1
    Gradient Picture2, Val(txtrgb(0)), Val(txtrgb(1)), Val(txtrgb(2)), True
End Sub

Private Sub cmdRand2_Click(Index As Integer)
    Randomize
    Select Case Index
        Case 0
            txtrgb(0).Text = Int(Rnd * 255) + 1
        Case 1
            txtrgb(1).Text = Int(Rnd * 255) + 1
        Case 2
            txtrgb(2).Text = Int(Rnd * 255) + 1
    End Select
     Gradient Picture2, Val(txtrgb(0)), Val(txtrgb(1)), Val(txtrgb(2)), True
End Sub

Private Sub cmdScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        isOn = True
    End If
End Sub

Private Sub cmdScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim mPos As POINTAPI
Dim mWnd As Long, mDc As Long
Dim stdout
    If Button = vbLeftButton And isOn Then
        GetCursorPos mPos
        KeepCol = GetPixel(GetDC(stdout), mPos.X, mPos.Y)
        LongToRgb (KeepCol)
        txtrgb(0).Text = T_RGB.Red
        txtrgb(1).Text = T_RGB.Green
        txtrgb(2).Text = T_RGB.Blue
    End If
End Sub

Private Sub cmdScreen_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Gradient Picture2, T_RGB.Red, T_RGB.Green, T_RGB.Blue, 1
    isOn = False
End Sub



Private Sub Form_Load()
Dim Count As Integer
    LoadPallet FixPath(App.Path) & "pallets\coolXp.pal"
    For Count = 0 To 32 - 1
        Picture1(Count).BackColor = Pallet(Count)
    Next
    FirstTime = True
    Count = 0
    
    lstOp.AddItem "C++ Hex"
    lstOp.AddItem "Delphi Hex"
    lstOp.AddItem "Visual Basic Hex"
    lstOp.AddItem "VB Long Colour"
    lstOp.AddItem "RGB Colour"
    lstOp.AddItem "HTML WebHex"
    lstOp.ListIndex = 0
    '
    lblrgb(0).ForeColor = vbRed
    lblrgb(1).ForeColor = vbGreen
    lblrgb(2).ForeColor = vbBlue
    
End Sub

Private Sub Form_Paint()
    Line3D3.Width = Form1.Width
    Line3D4.Width = Form1.Width
    Line3D5.Width = Form1.Width
    If Not FirstTime Then Exit Sub
    Picture1_Click 0
    FirstTime = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
    Set frmabout = Nothing
    End
End Sub

Private Sub hsbblue_Change()
    lblblue.Caption = hsbblue.Value
    colview.BackColor = RGB(Val(lblred), Val(lblgreen), Val(lblblue))
    'txthtm.Text = "#" & RGBtoHEX(colview.BackColor)
    txtrgb(0).Text = lblred
    txtrgb(1).Text = lblgreen
    txtrgb(2).Text = lblblue
    
End Sub

Private Sub hsbblue_Scroll()
    hsbblue_Change
End Sub

Private Sub hsbgreen_Change()
    lblgreen.Caption = hsbgreen.Value
    colview.BackColor = RGB(Val(lblred), Val(lblgreen), Val(lblblue))
    txtrgb(0).Text = lblred
    txtrgb(1).Text = lblgreen
    txtrgb(2).Text = lblblue
    'txthtm.Text = "#" & RGBtoHEX(colview.BackColor)
End Sub

Private Sub hsbgreen_Scroll()
    hsbgreen_Change
End Sub

Private Sub hsbred_Change()
    lblred.Caption = hsbred.Value
    colview.BackColor = RGB(Val(lblred), Val(lblgreen), Val(lblblue))
    
    'txthtm.Text = "#" & RGBtoHEX(colview.BackColor)
    txtrgb(0).Text = lblred
    txtrgb(1).Text = lblgreen
    txtrgb(2).Text = lblblue
End Sub

Private Sub hsbred_Scroll()
    hsbred_Change
End Sub

Private Sub lstOp_Click()
    LsIdx = lstOp.ListIndex
    GetColour
End Sub

Private Sub mnuabout_Click()
    frmabout.Show vbModal, Form1
End Sub

Private Sub mnuexit_Click()
Dim Ans As Integer, I As Integer
    Ans = MsgBox("Are you sure you would like to quit this program now?", vbYesNo Or vbQuestion, Form1.Caption)
    If Ans = vbNo Then Exit Sub
    StrRgb = ""
    KeepCol = 0
    nAction = 0
    LsIdx = 0
    Erase Pallet()
    
    For I = 0 To Picture1.Count - 1
        Set Picture1(I) = Nothing
    Next
    
    
    Unload Form1
    
End Sub

Private Sub mnunewpal_Click()
    frmnew.Show vbModal, Form1
End Sub

Private Sub mnuopen_Click()
Dim FileExt As String
Dim Count As Integer

    Cdialog.DialogTitle = "Open Pallet"
    Cdialog.Filter = "Pallet Files(*.pal)|*.pal"
    Cdialog.InitDir = FixPath(App.Path) & "pallets\"
    Cdialog.ShowOpen

    FileExt = Right(UCase(Cdialog.FileName), 3)
    If Len(Cdialog.FileName) <= 0 Then Exit Sub
    If Not FileExt = "PAL" Then
        MsgBox "This is not a vaild DM Pallet filename", vbInformation, Form1.Caption
        Exit Sub
    Else
        If LoadPallet(Cdialog.FileName) < 1 Then
            MsgBox "There was an error while loading the pallet file.", vbExclamation, Form1.Caption
            FileExt = ""
            Exit Sub
        Else
            For Count = 0 To 32 - 1
                Picture1(Count).BackColor = Pallet(Count)
            Next
            Count = 0
            Erase Pallet
        End If
    End If
    
End Sub

Private Sub Picture1_Click(Index As Integer)
    On Error Resume Next
    LongToRgb Picture1(Index).BackColor
    txtrgb(0).Text = T_RGB.Red
    txtrgb(1).Text = T_RGB.Green
    txtrgb(2).Text = T_RGB.Blue
    'txthtm.Text = "#" & RGBtoHEX(Picture1(Index).BackColor)
    colview.BackColor = Picture1(Index).BackColor

    hsbred.Value = T_RGB.Red
    hsbgreen.Value = T_RGB.Green
    hsbblue.Value = T_RGB.Blue
    
    Gradient Picture2, T_RGB.Red, T_RGB.Green, T_RGB.Blue, True
    GetColour
End Sub

Private Sub Picture2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim T_ColVal As Long
On Error Resume Next
    If Button = vbLeftButton Then
        T_ColVal = Picture2.Point(X, Y)
        LongToRgb T_ColVal
        colview.BackColor = T_ColVal
        hsbred.Value = T_RGB.Red
        hsbgreen.Value = T_RGB.Green
        hsbblue.Value = T_RGB.Blue
        GetColour
    End If
    ReleaseCapture

End Sub

Private Sub txtrgb_Change(Index As Integer)
On Error Resume Next
    If Val(txtrgb(Index).Text) > 255 Then txtrgb(Index).Text = 255
    hsbred.Value = Val(txtrgb(0).Text)
    hsbgreen.Value = Val(txtrgb(1).Text)
    hsbblue.Value = Val(txtrgb(2).Text)
    GetColour
    
End Sub
