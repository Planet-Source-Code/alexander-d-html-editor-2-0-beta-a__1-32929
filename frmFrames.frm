VERSION 5.00
Begin VB.Form frmFrames 
   Caption         =   "Frame Wizard"
   ClientHeight    =   4470
   ClientLeft      =   4605
   ClientTop       =   3390
   ClientWidth     =   8745
   Icon            =   "frmFrames.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4470
   ScaleWidth      =   8745
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "frmFrames.frx":030A
      Top             =   3960
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   28
      Text            =   "frmFrames.frx":0472
      Top             =   3960
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cle&ar"
      Height          =   375
      Left            =   2640
      TabIndex        =   27
      Top             =   3960
      Width           =   1335
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   26
      Text            =   "frmFrames.frx":05B9
      Top             =   3960
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   25
      Text            =   "frmFrames.frx":06F6
      Top             =   3960
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   24
      Text            =   "frmFrames.frx":08C6
      Top             =   3960
      Width           =   3135
   End
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00C0FFFF&
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   4155
      TabIndex        =   20
      Top             =   1200
      Width           =   4215
      Begin VB.Label Label12 
         BackColor       =   &H00CB9756&
         Height          =   1335
         Left            =   0
         TabIndex        =   22
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label11 
         BackColor       =   &H000000FF&
         Height          =   1215
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   4155
      TabIndex        =   17
      Top             =   1200
      Width           =   4215
      Begin VB.Label Label10 
         BackColor       =   &H000000FF&
         Height          =   2655
         Left            =   2160
         TabIndex        =   19
         Top             =   0
         Width           =   2055
      End
      Begin VB.Label Label9 
         BackColor       =   &H000000FF&
         Height          =   2655
         Left            =   0
         TabIndex        =   18
         Top             =   0
         Width           =   2055
      End
   End
   Begin VB.CommandButton cmdPlace 
      Caption         =   "&Place Code"
      Height          =   375
      Left            =   4080
      TabIndex        =   16
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate Code"
      Height          =   375
      Left            =   5640
      TabIndex        =   15
      Top             =   3960
      Width           =   1455
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   14
      Top             =   3960
      Width           =   1335
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   4155
      TabIndex        =   10
      Top             =   1200
      Width           =   4215
      Begin VB.Label Label8 
         BackColor       =   &H000000FF&
         Height          =   2055
         Left            =   960
         TabIndex        =   13
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label7 
         BackColor       =   &H000000FF&
         Height          =   495
         Left            =   960
         TabIndex        =   12
         Top             =   0
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackColor       =   &H000000FF&
         Height          =   2655
         Left            =   0
         TabIndex        =   11
         Top             =   0
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Code"
      Height          =   3255
      Left            =   4680
      TabIndex        =   8
      Top             =   600
      Width           =   3975
      Begin VB.TextBox txtCode 
         Height          =   2895
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   9
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label13 
         Caption         =   """"
         Height          =   615
         Left            =   720
         TabIndex        =   23
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   4155
      TabIndex        =   5
      Top             =   1200
      Width           =   4215
      Begin VB.Label Label3 
         BackColor       =   &H000000FF&
         Height          =   1935
         Left            =   0
         TabIndex        =   7
         Top             =   720
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackColor       =   &H000000FF&
         Height          =   615
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   4215
      End
   End
   Begin VB.ComboBox cboframe 
      Height          =   315
      ItemData        =   "frmFrames.frx":0A1E
      Left            =   480
      List            =   "frmFrames.frx":0A34
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   3735
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   240
      ScaleHeight     =   2595
      ScaleWidth      =   4155
      TabIndex        =   1
      Top             =   1200
      Width           =   4215
      Begin VB.Label Label5 
         BackColor       =   &H000000FF&
         Height          =   2655
         Left            =   840
         TabIndex        =   3
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Label4 
         BackColor       =   &H000000FF&
         Height          =   2655
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      Caption         =   "Select an Item on the list to view the preview"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   720
      TabIndex        =   30
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Use this to Generate your frames page. after you have the code, fill in the source for the pages and your ready to go!"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8655
   End
End
Attribute VB_Name = "frmFrames"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboframe_Click()
If cboframe.Text = "Contents and Main" Then
Picture1.Visible = True
Picture2.Visible = False
Picture3.Visible = False
Picture5.Visible = False
Picture4.Visible = False
ElseIf cboframe.Text = "Top and Bottom" Then
Picture2.Visible = True
Picture1.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
ElseIf cboframe.Text = "Contents, Top and Bottom" Then
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = True
Picture4.Visible = False
Picture5.Visible = False
ElseIf cboframe.Text = "Half and Half Vertical" Then
Picture4.Visible = True
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture5.Visible = False
ElseIf cboframe.Text = "Half and Half Horizontal" Then
Picture5.Visible = True
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Else
Picture5.Visible = False
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
End If
End Sub

Private Sub cmdCAncel_Click()
Unload Me
End Sub

Private Sub cmdGenerate_Click()
Dim strQ As String
strQ = Label13.Caption
If Picture1.Visible = True Then
txtCode.Text = Text1.Text
ElseIf Picture3.Visible = True Then
txtCode.Text = Text2.Text
ElseIf Picture5.Visible = True Then
txtCode.Text = Text3.Text
ElseIf Picture4.Visible = True Then
txtCode.Text = Text4.Text
ElseIf Picture2.Visible = True Then
txtCode.Text = Text5.Text
Else
MsgBox "Please select an item from the list.", vbExclamation, "Select...."
cboframe.SetFocus
End If
End Sub

Private Sub cmdPlace_Click()
Form1.Label13.Caption = "GO"
Form1.Text1.Text = txtCode.Text
Unload Me
Form1.Show
End Sub

Private Sub Command1_Click()
txtCode.Text = ""
End Sub

Private Sub Form_Load()
cboframe.ListIndex = 0
Text4.Visible = False
Text1.Visible = False
Text3.Visible = False
Text2.Visible = False
Text5.Visible = False
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Label2.BackColor = &HCB9756
Label3.BackColor = &HCB9756
Label4.BackColor = &HCB9756
Label5.BackColor = &HCB9756
Label6.BackColor = &HCB9756
Label7.BackColor = &HCB9756
Label8.BackColor = &HCB9756
Label9.BackColor = &HCB9756
Label10.BackColor = &HCB9756
Label11.BackColor = &HCB9756
Label12.BackColor = &HCB9756
Picture1.BackColor = &HC0FFFF
Picture2.BackColor = &HC0FFFF
Picture3.BackColor = &HC0FFFF
Picture4.BackColor = &HC0FFFF
Picture5.BackColor = &HC0FFFF
End Sub

