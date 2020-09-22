VERSION 5.00
Begin VB.Form frmStyle 
   Caption         =   "Link Style For your page..."
   ClientHeight    =   4830
   ClientLeft      =   5820
   ClientTop       =   4095
   ClientWidth     =   6780
   Icon            =   "frmStyle.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4830
   ScaleWidth      =   6780
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Code"
      Height          =   375
      Left            =   2640
      TabIndex        =   14
      Top             =   4320
      Width           =   1215
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   3960
      TabIndex        =   13
      Top             =   4320
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Close"
      Height          =   375
      Left            =   5400
      TabIndex        =   5
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Your Code"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   6495
      Begin VB.TextBox txtStyleCode 
         Height          =   1695
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   480
         Width           =   6255
      End
      Begin VB.Label Label1 
         Caption         =   "Place the code in the header of the page you want it to effect."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   615
         Left            =   240
         TabIndex        =   3
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "A link when the mouse is over it..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   6495
      Begin VB.ComboBox cboHTrans 
         Height          =   315
         ItemData        =   "frmStyle.frx":030A
         Left            =   3960
         List            =   "frmStyle.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   1815
      End
      Begin VB.CheckBox chkHUnderline 
         Caption         =   "Underline"
         Height          =   255
         Left            =   2280
         TabIndex        =   11
         Top             =   360
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtHcolor 
         Height          =   285
         Left            =   600
         TabIndex        =   10
         Text            =   "Red"
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Color"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Just a regular link on the page..."
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      Begin VB.CheckBox chkUnderline 
         Caption         =   "Underlined"
         Height          =   255
         Left            =   2280
         TabIndex        =   8
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox txtColor 
         Height          =   285
         Left            =   600
         TabIndex        =   7
         Text            =   "Blue"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCAncel_Click()
Unload Me
End Sub

Private Sub cmdGenerate_Click()
Dim strQ As String
Dim strColor As String
Dim strunderline As String
Dim strHcolor As String
Dim strHunderline As String
Dim strHTrans As String

If chkUnderline.Value = Checked Then
strunderline = "underline;}"
Else
strunderline = "none;}"
End If

If chkHUnderline.Value = Checked Then
strHunderline = "underline; "
Else
strHunderline = "none; "
End If

strColor = txtColor.Text & "; "

strQ = ""
strHcolor = txtHcolor.Text & "; "

If cboHTrans.Text = "Capitalize" Then
strHTrans = " capitalize;}"
ElseIf cboHTrans.Text = "Make Uppercase" Then
strHTrans = " uppercase;}"
ElseIf cboHTrans.Text = "Make Lowercase" Then
strHTrans = " lowercase;}"
Else
strHTrans = " none;)"
End If

tag1 = "<style>"
tag2 = "<!--"
tag3 = "a       {color: " & strColor & "text-decoration: " & strunderline
tag4 = "a:hover {color: " & strHcolor & "text-decoration: " & strHunderline & "text-transform: " & strHTrans
tag5 = "-->"
tag6 = "</style>"

txtStyleCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6
Form1.Label13.Caption = "GO"
End Sub

Private Sub Command1_Click()
Clipboard.SetText (txtStyleCode.Text)
End Sub

Private Sub Form_Load()
cboHTrans.ListIndex = 0
End Sub
