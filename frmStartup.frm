VERSION 5.00
Begin VB.Form frmStartup 
   Caption         =   "Startup..."
   ClientHeight    =   4215
   ClientLeft      =   6825
   ClientTop       =   3690
   ClientWidth     =   5055
   Icon            =   "frmStartup.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4215
   ScaleWidth      =   5055
   Begin VB.CommandButton Command1 
      Caption         =   "Open Existing..."
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   3720
      Width           =   1695
   End
   Begin VB.VScrollBar VScroll1 
      Enabled         =   0   'False
      Height          =   3015
      Left            =   4680
      TabIndex        =   8
      Top             =   600
      Width           =   255
   End
   Begin VB.CommandButton cmdContinue 
      Caption         =   "&New..."
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3720
      Width           =   1335
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   3360
      TabIndex        =   3
      Top             =   3720
      Width           =   1575
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      ScaleHeight     =   2955
      ScaleWidth      =   4515
      TabIndex        =   0
      Top             =   600
      Width           =   4575
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Blank Page"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Image Image4 
         Height          =   855
         Left            =   120
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Script"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   7
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image Image3 
         Height          =   930
         Left            =   3240
         Picture         =   "frmStartup.frx":030A
         Top             =   120
         Width           =   1020
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Frames Wizard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1680
         TabIndex        =   6
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image Image2 
         Height          =   930
         Left            =   1680
         Picture         =   "frmStartup.frx":0E06
         Top             =   120
         Width           =   1020
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00C00000&
         Caption         =   "Page Wizard"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Image Image1 
         Height          =   930
         Left            =   120
         Picture         =   "frmStartup.frx":1902
         Top             =   120
         Width           =   1020
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Select an Item to continue"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3975
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   15
   End
End
Attribute VB_Name = "frmStartup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdContinue_Click()
If Label3.BackColor = &HC00000 Then
frmMain.Show
ElseIf Label4.BackColor = &HC00000 Then
frmFrames.Show
ElseIf Label6.BackColor = &HC00000 Then
Form1.Show
Else
frmScript.Show
End If
Unload Me
End Sub

Private Sub cmdExit_Click()
End
End Sub

Private Sub Command1_Click()
FileOpenProc
Unload Me
Form1.Show
End Sub

Private Sub Form_Load()
Image4.Picture = Image1.Picture
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Label1.Caption = "GO"
End Sub

Private Sub Image1_Click()
Label3.BackColor = &HC00000
Label3.ForeColor = vbWhite
Label4.BackColor = vbWhite
Label4.ForeColor = vbBlack
Label5.BackColor = vbWhite
Label5.ForeColor = vbBlack
Label6.BackColor = vbWhite
Label6.ForeColor = vbBlack
End Sub

Private Sub Image1_DblClick()
frmMain.Show
Unload Me
End Sub

Private Sub Image2_Click()
Label3.BackColor = vbWhite
Label3.ForeColor = vbBlack
Label4.BackColor = &HC00000
Label4.ForeColor = vbWhite
Label5.BackColor = vbWhite
Label5.ForeColor = vbBlack
Label6.BackColor = vbWhite
Label6.ForeColor = vbBlack
End Sub

Private Sub Image2_DblClick()
frmFrames.Show
Unload Me
End Sub

Private Sub Image3_Click()
Label6.BackColor = vbWhite
Label6.ForeColor = vbBlack
Label5.BackColor = &HC00000
Label5.ForeColor = vbWhite
Label3.BackColor = vbWhite
Label3.ForeColor = vbBlack
Label4.BackColor = vbWhite
Label4.ForeColor = vbBlack
End Sub

Private Sub Image3_DblClick()
frmScript.Show
Unload Me
End Sub

Private Sub Image4_Click()
Label4.BackColor = vbWhite
Label4.ForeColor = vbBlack
Label5.BackColor = vbWhite
Label5.ForeColor = vbBlack
Label3.BackColor = vbWhite
Label3.ForeColor = vbBlack
Label6.BackColor = &HC00000
Label6.ForeColor = vbWhite
End Sub

Private Sub Image4_DblClick()
Form1.Show
Unload Me
End Sub

Private Sub Label3_Click()
Label3.BackColor = &HC00000
Label3.ForeColor = vbWhite
Label4.BackColor = vbWhite
Label4.ForeColor = vbBlack
Label5.BackColor = vbWhite
Label5.ForeColor = vbBlack
Label6.BackColor = vbWhite
Label6.ForeColor = vbBlack
Unload Me
End Sub

Private Sub Label3_DblClick()
frmMain.Show
Unload Me
End Sub

Private Sub Label4_Click()
Label6.BackColor = vbWhite
Label6.ForeColor = vbBlack
Label3.BackColor = vbWhite
Label3.ForeColor = vbBlack
Label4.BackColor = &HC00000
Label4.ForeColor = vbWhite
Label5.BackColor = vbWhite
Label5.ForeColor = vbBlack
Unload Me
End Sub

Private Sub Label4_DblClick()
frmFrames.Show
Unload Me
End Sub

Private Sub Label5_Click()
Label6.BackColor = vbWhite
Label6.ForeColor = vbBlack
Label5.BackColor = &HC00000
Label5.ForeColor = vbWhite
Label3.BackColor = vbWhite
Label3.ForeColor = vbBlack
Label4.BackColor = vbWhite
Label4.ForeColor = vbBlack
Unload Me
End Sub

Private Sub Label5_DblClick()
frmScript.Show
Unload Me
End Sub

Private Sub Label6_Click()
Label4.BackColor = vbWhite
Label4.ForeColor = vbBlack
Label5.BackColor = vbWhite
Label5.ForeColor = vbBlack
Label3.BackColor = vbWhite
Label3.ForeColor = vbBlack
Label6.BackColor = &HC00000
Label6.ForeColor = vbWhite
End Sub

Private Sub Label6_DblClick()
Form1.Show
Unload Me
End Sub

