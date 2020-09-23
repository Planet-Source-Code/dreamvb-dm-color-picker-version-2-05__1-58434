VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmnew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create new Pallet"
   ClientHeight    =   2520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2700
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2520
   ScaleWidth      =   2700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin Project1.Flat2 Flat21 
      Height          =   1740
      Left            =   105
      TabIndex        =   35
      Top             =   60
      Width           =   2460
      _ExtentX        =   4339
      _ExtentY        =   3069
   End
   Begin VB.CommandButton cmdclose 
      Caption         =   "&Close"
      Height          =   345
      Left            =   1380
      TabIndex        =   34
      Top             =   1965
      Width           =   1020
   End
   Begin VB.CommandButton cmdsave 
      Caption         =   "&Save"
      Height          =   345
      Left            =   195
      TabIndex        =   33
      Top             =   1965
      Width           =   1020
   End
   Begin MSComDlg.CommonDialog Cdialog 
      Left            =   2730
      Top             =   510
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   0
      Left            =   450
      MouseIcon       =   "frmnew.frx":0000
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   31
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   1
      Left            =   660
      MouseIcon       =   "frmnew.frx":030A
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   30
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   2
      Left            =   870
      MouseIcon       =   "frmnew.frx":0614
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   29
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   3
      Left            =   1080
      MouseIcon       =   "frmnew.frx":091E
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   28
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   4
      Left            =   1290
      MouseIcon       =   "frmnew.frx":0C28
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   27
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   5
      Left            =   1500
      MouseIcon       =   "frmnew.frx":0F32
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   26
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   6
      Left            =   1710
      MouseIcon       =   "frmnew.frx":123C
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   25
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   7
      Left            =   1920
      MouseIcon       =   "frmnew.frx":1546
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   24
      Top             =   660
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   8
      Left            =   450
      MouseIcon       =   "frmnew.frx":1850
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   23
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   9
      Left            =   660
      MouseIcon       =   "frmnew.frx":1B5A
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   22
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   10
      Left            =   870
      MouseIcon       =   "frmnew.frx":1E64
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   21
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   11
      Left            =   1080
      MouseIcon       =   "frmnew.frx":216E
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   20
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   12
      Left            =   1290
      MouseIcon       =   "frmnew.frx":2478
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   19
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   13
      Left            =   1500
      MouseIcon       =   "frmnew.frx":2782
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   18
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   14
      Left            =   1710
      MouseIcon       =   "frmnew.frx":2A8C
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   17
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   15
      Left            =   1920
      MouseIcon       =   "frmnew.frx":2D96
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   16
      Top             =   870
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   16
      Left            =   450
      MouseIcon       =   "frmnew.frx":30A0
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   15
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   17
      Left            =   660
      MouseIcon       =   "frmnew.frx":33AA
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   14
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   18
      Left            =   870
      MouseIcon       =   "frmnew.frx":36B4
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   13
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   19
      Left            =   1080
      MouseIcon       =   "frmnew.frx":39BE
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   12
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   20
      Left            =   1290
      MouseIcon       =   "frmnew.frx":3CC8
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   11
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   21
      Left            =   1500
      MouseIcon       =   "frmnew.frx":3FD2
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   10
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   22
      Left            =   1710
      MouseIcon       =   "frmnew.frx":42DC
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   9
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   23
      Left            =   1920
      MouseIcon       =   "frmnew.frx":45E6
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   8
      Top             =   1080
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   24
      Left            =   450
      MouseIcon       =   "frmnew.frx":48F0
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   7
      Top             =   1290
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   25
      Left            =   660
      MouseIcon       =   "frmnew.frx":4BFA
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   6
      Top             =   1290
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   26
      Left            =   870
      MouseIcon       =   "frmnew.frx":4F04
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   5
      Top             =   1290
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   27
      Left            =   1080
      MouseIcon       =   "frmnew.frx":520E
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   4
      Top             =   1290
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   28
      Left            =   1290
      MouseIcon       =   "frmnew.frx":5518
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   3
      Top             =   1290
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   29
      Left            =   1500
      MouseIcon       =   "frmnew.frx":5822
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   2
      Top             =   1290
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   30
      Left            =   1710
      MouseIcon       =   "frmnew.frx":5B2C
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   1
      Top             =   1290
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H80000008&
      Height          =   225
      Index           =   31
      Left            =   1920
      MouseIcon       =   "frmnew.frx":5E36
      ScaleHeight     =   195
      ScaleWidth      =   225
      TabIndex        =   0
      Top             =   1290
      Width           =   255
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Double click each box to add a new color."
      Height          =   435
      Left            =   375
      TabIndex        =   32
      Top             =   165
      Width           =   2025
   End
End
Attribute VB_Name = "frmnew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub cmdclose_Click()
    Unload frmnew
End Sub

Private Sub cmdsave_Click()
Dim StrBuff As String, FileNum As Long, I As Integer
Dim Ans
    ' this will build our pallet
    For I = 0 To Picture1.Count - 1
        LongToRgb Picture1(I).BackColor
        StrBuff = StrBuff & Chr(T_RGB.Red) & Chr(T_RGB.Green) & Chr(T_RGB.Blue) & ","
    Next
    I = 0
    
    CDialog.DialogTitle = "Save DM Pallet file As"
    CDialog.Filter = "DM Pallet Files(*.pal)|*.pal"
    CDialog.ShowSave
    If Len(Trim(CDialog.FileName)) <= 0 Then Exit Sub
    
    If FindFile(CDialog.FileName) = True Then
        Ans = MsgBox("The file already exists do you whish to replace this file", vbYesNo Or vbQuestion, frmnew.Caption)
        If Ans = vbNo Then
            Exit Sub
    Else
        SavePallet CDialog.FileName, StrBuff
        MsgBox "File saved to " & CDialog.FileName, vbInformation, Form1.Caption
        StrBuff = ""
        End If
        Exit Sub
    Else
        SavePallet CDialog.FileName, StrBuff
        MsgBox "File saved to " & CDialog.FileName, vbInformation, Form1.Caption
        StrBuff = ""
    End If

End Sub

Private Sub Form_Load()
Dim I As Integer
    For I = 0 To Picture1.Count - 1
        Picture1(I).BackColor = RGB(255, 0, 255)
    Next
    frmnew.Icon = Nothing
    I = 0
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmnew = Nothing
    Unload frmnew
    
    
End Sub

Private Sub Picture1_DblClick(Index As Integer)
    CDialog.ShowColor
    Picture1(Index).BackColor = CDialog.Color
    
End Sub
