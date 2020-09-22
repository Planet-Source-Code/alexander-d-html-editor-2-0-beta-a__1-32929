VERSION 5.00
Begin VB.Form frmHR 
   Caption         =   "Horizontal Lines"
   ClientHeight    =   3975
   ClientLeft      =   7035
   ClientTop       =   3690
   ClientWidth     =   5640
   Icon            =   "frmHR.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3975
   ScaleWidth      =   5640
   Begin VB.CommandButton cmdCopy 
      Caption         =   "C&opy Code"
      Height          =   375
      Left            =   1440
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   2880
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdCAncel 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4200
      TabIndex        =   9
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Your Code:"
      Height          =   1935
      Left            =   120
      TabIndex        =   7
      Top             =   1440
      Width           =   5415
      Begin VB.TextBox txtCode 
         Height          =   1575
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   240
         Width           =   5055
      End
      Begin VB.Label Label5 
         Caption         =   """"
         Height          =   615
         Left            =   1440
         TabIndex        =   14
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "HR Information"
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.ComboBox cboAlign 
         Height          =   315
         ItemData        =   "frmHR.frx":030A
         Left            =   3600
         List            =   "frmHR.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtColor 
         Height          =   285
         Left            =   960
         TabIndex        =   6
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   3000
         TabIndex        =   4
         Text            =   "90"
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox txtPixel 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Text            =   "1"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Alignment:"
         Height          =   255
         Left            =   2640
         TabIndex        =   13
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "Line Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Width (%):"
         Height          =   255
         Left            =   2160
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Pixels High:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmHR"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub cmdCAncel_Click()
Unload Me
End Sub

Private Sub cmdCopy_Click()
Clipboard.SetText (txtCode.Text)
End Sub

Private Sub cmdGen_Click()
Dim strQ As String
Dim strColor As String
Dim strHeight As String
Dim strAlign As String
Dim strwidth As String

strQ = Label5.Caption
strHeight = txtPixel.Text
strColor = txtColor.Text
strAlign = cboAlign.Text
strwidth = txtWidth.Text

Tag = "<hr size= " & strQ & strHeight & strQ & " color = " & strQ & strColor & strQ & " align = " & strQ & strAlign & strQ & " width = " & strQ & strwidth & "%" & strQ & " >"

txtCode.Text = Tag
Form1.Label13.Caption = "GO"
End Sub

Private Sub Form_Load()
cboAlign.ListIndex = 0
End Sub
