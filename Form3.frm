VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00C0C0C0&
   Caption         =   "HTML Editor Options"
   ClientHeight    =   3975
   ClientLeft      =   6225
   ClientTop       =   3795
   ClientWidth     =   5595
   LinkTopic       =   "Form3"
   ScaleHeight     =   3975
   ScaleWidth      =   5595
   Begin VB.CommandButton Command14 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   4320
      TabIndex        =   22
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command13 
      Caption         =   "&Apply"
      Height          =   375
      Left            =   3120
      TabIndex        =   21
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton Command12 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1920
      TabIndex        =   20
      Top             =   3480
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   5953
      _Version        =   327680
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Color Schemes"
      TabPicture(0)   =   "Form3.frx":0000
      Tab(0).ControlCount=   7
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Option6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Option4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Option1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Option2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Option3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Command4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Option5"
      Tab(0).Control(6).Enabled=   0   'False
      TabCaption(1)   =   "Misc."
      TabPicture(1)   =   "Form3.frx":001C
      Tab(1).ControlCount=   2
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(1).Enabled=   0   'False
      TabCaption(2)   =   "About"
      TabPicture(2)   =   "Form3.frx":0038
      Tab(2).ControlCount=   2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture1"
      Tab(2).Control(0).Enabled=   -1  'True
      Tab(2).Control(1)=   "Label2"
      Tab(2).Control(1).Enabled=   0   'False
      Begin VB.OptionButton Option5 
         Caption         =   "White Backround and Red Foreground"
         Height          =   375
         Left            =   600
         TabIndex        =   25
         Top             =   1440
         Width           =   3135
      End
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   -74880
         Picture         =   "Form3.frx":0054
         ScaleHeight     =   1815
         ScaleWidth      =   5415
         TabIndex        =   14
         Top             =   480
         Width           =   5415
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Version 1.0"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   3240
            TabIndex        =   15
            Top             =   1320
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "HTML"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00A26C1E&
            Height          =   495
            Left            =   2640
            TabIndex        =   16
            Top             =   1080
            Width           =   2535
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Font"
         Height          =   1455
         Left            =   -74640
         TabIndex        =   8
         Top             =   1320
         Width           =   4815
         Begin VB.CommandButton cmdDown 
            Caption         =   "DOWN"
            Height          =   195
            Left            =   3960
            TabIndex        =   24
            Top             =   1080
            Width           =   735
         End
         Begin VB.CommandButton cmdUp 
            Caption         =   "UP"
            Height          =   195
            Left            =   3960
            TabIndex        =   23
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton Command11 
            Caption         =   "Command11"
            Height          =   255
            Left            =   840
            TabIndex        =   19
            Top             =   1080
            Width           =   495
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   3120
            MaxLength       =   2
            TabIndex        =   18
            Text            =   "10"
            Top             =   840
            Width           =   735
         End
         Begin VB.CommandButton Command9 
            Caption         =   "Command9"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   1080
            Width           =   495
         End
         Begin VB.CommandButton Command8 
            Caption         =   "Command8"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1080
            Width           =   495
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Form3.frx":1FFB
            Left            =   120
            List            =   "Form3.frx":201A
            TabIndex        =   10
            Text            =   "Tahoma"
            Top             =   240
            Width           =   3975
         End
         Begin VB.Label Label5 
            Caption         =   "Font Size:"
            Height          =   255
            Left            =   2280
            TabIndex        =   17
            Top             =   840
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "Tahoma"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   9
            Top             =   840
            Width           =   4215
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Display At Startup"
         Height          =   735
         Left            =   -74640
         TabIndex        =   6
         Top             =   480
         Width           =   4815
         Begin VB.CheckBox Check2 
            Caption         =   "File Browser"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   240
            Value           =   1  'Checked
            Width           =   2295
         End
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Command4"
         Height          =   375
         Left            =   4560
         TabIndex        =   5
         Top             =   1200
         Width           =   375
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Black Backround and White Foreground"
         Height          =   375
         Left            =   600
         TabIndex        =   3
         Top             =   840
         Width           =   3375
      End
      Begin VB.OptionButton Option2 
         Caption         =   "White Backround and Black Foreground"
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   2040
         Width           =   3255
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Black Backround and Gray Foreground"
         Height          =   255
         Left            =   600
         TabIndex        =   1
         Top             =   600
         Value           =   -1  'True
         Width           =   3975
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Black Backround and Red Foreground"
         Height          =   495
         Left            =   600
         TabIndex        =   4
         Top             =   1080
         Width           =   3255
      End
      Begin VB.OptionButton Option6 
         Caption         =   "White Backround and Green Foreground"
         Height          =   495
         Left            =   600
         TabIndex        =   26
         Top             =   1680
         Width           =   3495
      End
      Begin VB.Label Label2 
         Caption         =   "Davicsoft HTML Editor was made by Alex Davis."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -74520
         TabIndex        =   13
         Top             =   2400
         Width           =   4695
      End
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDown_Click()
Text1.Text = Text1.Text - 1
End Sub

Private Sub cmdUp_Click()
Text1.Text = Text1.Text + 1
End Sub

Private Sub Combo1_Click()
Label1.Caption = Combo1.Text
Label1.Font = Combo1.Text
End Sub

Private Sub Command1_Click()
If Option1.Value = True Then
Form1.Text2.Text = "OPT1"
Form1.Text1.BackColor = vbBlack
Form1.Text1.ForeColor = &HC0C0C0
Else
If Option2.Value = True Then
Form1.Text2.Text = "OPT2"
Form1.Text1.BackColor = vbWhite
Form1.Text1.ForeColor = vbBlack
Else
If Option3.Value = True Then
Form1.Text2.Text = "OPT3"
Form1.Text1.BackColor = vbBlack
Form1.Text1.ForeColor = vbWhite
Else
If Option4.Value = True Then
Form1.Text2.Text = "OPT4"
Form1.Text1.BackColor = vbBlack
Form1.Text1.ForeColor = vbRed
End If
End If
End If
End If
Command4_Click
End Sub

Private Sub Command10_Click()
Unload Me
End Sub

Private Sub Command11_Click()
Open App.Path & "\option4.txt" For Output As #1
    Print #1, Form1.Text5.Text
Close
End Sub

Private Sub Command12_Click()
Command5_Click
Command8_Click
Command11_Click
Command1_Click
Command4_Click
Unload Me
End Sub

Private Sub Command13_Click()
If Option1.Value = True Then
Form1.Text2.Text = "OPT1"
Form1.Text1.BackColor = vbBlack
Form1.Text1.ForeColor = &HC0C0C0
Else
If Option2.Value = True Then
Form1.Text2.Text = "OPT2"
Form1.Text1.BackColor = vbWhite
Form1.Text1.ForeColor = vbBlack
Else
If Option3.Value = True Then
Form1.Text2.Text = "OPT3"
Form1.Text1.BackColor = vbBlack
Form1.Text1.ForeColor = vbWhite
Else
If Option4.Value = True Then
Form1.Text2.Text = "OPT4"
Form1.Text1.BackColor = vbBlack
Form1.Text1.ForeColor = vbRed
Else
If Option5.Value = True Then
Form1.Text2.Text = "OPT5"
Form1.Text1.BackColor = vbWhite
Form1.Text1.ForeColor = vbRed
Else
If Option6.Value = True Then
Form1.Text2.Text = "OPT6"
Form1.Text1.BackColor = vbWhite
Form1.Text1.ForeColor = &H8000&
End If
End If
End If
End If
End If
End If
Command4_Click

Dim intnumber As Integer
intnumber = Text1.Text

If Text1.Text = "" Then
Else
If IsNumeric(intnumber) = False Or intnumber >= 21 Or intnumber <= 9 Then
MsgBox "Please enter a valid number between 10 and 20", vbCritical, "Error"
Text1.Text = ""
Text1.SetFocus
Else
Form1.Text1.FontSize = intnumber
Form1.Text5.Text = intnumber
End If
End If

Form1.Text3.Text = Combo1.Text
Form1.Text1.Font = Combo1.Text


If Check2.Value = 1 Then
Form1.fbdick.Checked = True
Form1.Dir1.Visible = True
Form1.Label2.Visible = True
Form1.Drive1.Visible = True
Form1.File1.Visible = True
Form1.Height = Form1.Height + 10
Form1.Text4.Text = "FB"

Else
If Check2.Value = 0 Then
Form1.fbdick.Checked = False
Form1.Dir1.Visible = False
Form1.Label2.Visible = False
Form1.Drive1.Visible = False
Form1.File1.Visible = False
Form1.Height = Form1.Height - 10
Form1.Text4.Text = ""
End If
End If
Command8_Click
Command9_Click
Command11_Click
End Sub

Private Sub Command14_Click()
Unload Me
End Sub

Private Sub Command2_Click()
Command1_Click
Command4_Click
Form1.Text1.FontSize = 10
Unload Me
End Sub

Private Sub Command3_Click()
Unload Me
End Sub

Private Sub Command4_Click()
Open App.Path & "\option1.txt" For Output As #1
    Print #1, Form1.Text2.Text
Close
End Sub

Private Sub Command5_Click()
Dim intnumber As Integer
intnumber = Text1.Text

If Text1.Text = "" Then
Else
If IsNumeric(intnumber) = False Or intnumber >= 21 Or intnumber <= 9 Then
MsgBox "Please enter a valid number between 10 and 20", vbCritical, "Error"
Text1.Text = ""
Text1.SetFocus
Else
Form1.Text1.FontSize = intnumber
Form1.Text5.Text = intnumber
End If
End If

Form1.Text3.Text = Combo1.Text
Form1.Text1.Font = Combo1.Text


If Check2.Value = 1 Then
Form1.fbdick.Checked = True
Form1.Dir1.Visible = True
Form1.Label2.Visible = True
Form1.Drive1.Visible = True
Form1.File1.Visible = True
Form1.Height = Form1.Height + 10
Form1.Text4.Text = "FB"

Else
If Check2.Value = 0 Then
Form1.fbdick.Checked = False
Form1.Dir1.Visible = False
Form1.Label2.Visible = False
Form1.Drive1.Visible = False
Form1.File1.Visible = False
Form1.Height = Form1.Height - 10
Form1.Text4.Text = ""
End If
End If
Command8_Click
Command9_Click
Command11_Click
End Sub

Private Sub Command6_Click()
Unload Me
End Sub

Private Sub Command7_Click()
Command5_Click
Command8_Click
Command11_Click
Unload Me
End Sub

Private Sub Command8_Click()
Open App.Path & "\option2.txt" For Output As #1
    Print #1, Form1.Text3.Text
Close
End Sub

Private Sub Command9_Click()
Open App.Path & "\option3.txt" For Output As #1
    Print #1, Form1.Text4.Text
Close
End Sub

Private Sub Form_Load()
Text1.Text = Form1.Text5.Text
Combo1.Text = Form1.Text1.FontName
If Form1.Text2.Text = "OPT1" Then
Option1.Value = True
Else
If Form1.Text2.Text = "OPT2" Then
Option2.Value = True
Else
If Form1.Text2.Text = "OPT3" Then
Option3.Value = True
Else
If Form1.Text2.Text = "OPT4" Then
Option4.Value = True
Else
If Form1.Text2.Text = "OPT5" Then
Option5.Value = True
Else
If Form1.Text2.Text = "OPT6" Then
Option6.Value = True
End If
End If
End If
End If
End If
End If
Command11.Visible = False
Command9.Visible = False
Command8.Visible = False
Label1.Font = "Tahoma"
Command4.Visible = False
End Sub

Private Sub List1_Click()
List1.List = Label1.Caption
End Sub

