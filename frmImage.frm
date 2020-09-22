VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmImage 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Insert an Image"
   ClientHeight    =   5130
   ClientLeft      =   4305
   ClientTop       =   2985
   ClientWidth     =   10170
   Icon            =   "frmImage.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   5130
   ScaleWidth      =   10170
   Begin VB.CommandButton Command2 
      Caption         =   "Copy Code"
      Height          =   375
      Left            =   5160
      TabIndex        =   17
      Top             =   4680
      Width           =   1935
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4800
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   7200
      TabIndex        =   13
      Top             =   4680
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Caption         =   "Image Preview"
      Height          =   4335
      Left            =   5760
      TabIndex        =   11
      Top             =   240
      Width           =   4335
      Begin VB.PictureBox picImage 
         Height          =   3495
         Left            =   120
         ScaleHeight     =   3435
         ScaleWidth      =   4035
         TabIndex        =   12
         ToolTipText     =   "Picture Preview"
         Top             =   240
         Width           =   4095
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Image Preview only available if image is local. Please make sure you include ""http://"" in your locations and links."
            Height          =   1935
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   3855
         End
      End
      Begin VB.Label lblHeight 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   2640
         TabIndex        =   21
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label8 
         Caption         =   "Height:"
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   3960
         Width           =   735
      End
      Begin VB.Label lblWidth 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   720
         TabIndex        =   19
         Top             =   3960
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   3960
         Width           =   735
      End
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   9000
      TabIndex        =   9
      Top             =   4680
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "&Your Code"
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   5415
      Begin VB.TextBox txtCode 
         Height          =   1815
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   8
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label Label3 
         Caption         =   "Put this code in the body where you want the image to appear."
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   5055
      End
   End
   Begin VB.TextBox txtLink 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   1080
      Width           =   3375
   End
   Begin VB.CheckBox chkLink 
      Caption         =   "Make Link"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Input"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
      Begin VB.CommandButton Command1 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4800
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
      Begin VB.TextBox txtAlt 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox txtLocation 
         Height          =   285
         Left            =   1320
         TabIndex        =   2
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label2 
         Caption         =   "Tool Tip Text:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Image Location:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Label Label4 
      Caption         =   """"
      Height          =   495
      Left            =   6600
      TabIndex        =   14
      Top             =   3480
      Width           =   855
   End
End
Attribute VB_Name = "frmImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdGenerate_Click()
Dim strLoc As String
Dim strQ As String
Dim strAlt As String
Dim strHTM As String

If chkLink.Value = Checked Then
strLoc = txtLocation.Text
strQ = Label4.Caption
strAlt = txtAlt.Text
strHTM = txtLink.Text

tag1 = "<p><a href=" & strQ & strHTM & strQ & "><img src=" & strQ & strLoc & strQ & " alt =" & strQ & strAlt & strQ & "></a></p>"
txtCode.Text = tag1
Else
strLoc = txtLocation.Text
strQ = Label4.Caption
strAlt = txtAlt.Text

tag1 = "<p><img src=" & strQ & strLoc & strQ & " alt =" & strQ & strAlt & strQ & "></p>"
txtCode.Text = tag1
Form1.Label13.Caption = "GO"
End If
End Sub

Private Sub Command1_Click()
 Dim sFile As String
    With CommonDialog1
        .DialogTitle = "Open"
        .CancelError = False
        .Filter = "All Files (*.*)|*.*"
        .ShowOpen
        If Len(.filename) = 0 Then
            Exit Sub
        End If
        sFile = .filename
    End With
txtLocation.Text = sFile
End Sub

Private Sub Command2_Click()
Clipboard.SetText (txtCode.Text)
End Sub

Private Sub txtLocation_Change()
On Error Resume Next
picImage.Picture = LoadPicture(txtLocation.Text)
Label5.Visible = False
 lblHeight.Caption = picImage.Picture.Height
 lblWidth.Caption = picImage.Picture.Width
End Sub
