VERSION 5.00
Begin VB.Form frmScript 
   Caption         =   "Insert Script"
   ClientHeight    =   4425
   ClientLeft      =   5520
   ClientTop       =   3075
   ClientWidth     =   9270
   Icon            =   "frmScript.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4425
   ScaleWidth      =   9270
   Begin VB.CommandButton Command2 
      Caption         =   "Copy Code"
      Height          =   375
      Left            =   3480
      TabIndex        =   14
      Top             =   3960
      Width           =   1095
   End
   Begin VB.TextBox txtOther 
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Top             =   1080
      Width           =   2415
   End
   Begin VB.OptionButton optOther 
      Caption         =   "Other:"
      Height          =   375
      Left            =   360
      TabIndex        =   12
      Top             =   960
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Cl&ear"
      Height          =   375
      Left            =   6120
      TabIndex        =   9
      Top             =   3960
      Width           =   1575
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   3960
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   7800
      TabIndex        =   6
      Top             =   3960
      Width           =   1095
   End
   Begin VB.Frame Frame3 
      Caption         =   "Your Code"
      Height          =   3735
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   4215
      Begin VB.TextBox txtScriptCode 
         Height          =   3375
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   7
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label1 
         Caption         =   """"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Your Script"
      Height          =   2295
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   4455
      Begin VB.TextBox txtScript 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A26C1E&
         Height          =   2055
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   4
         Top             =   240
         Width           =   4095
      End
   End
   Begin VB.OptionButton optJava 
      Caption         =   "Java Script"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
   Begin VB.OptionButton optVBScript 
      Caption         =   "Visual Basic Script"
      Height          =   375
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Type Of Script"
      Height          =   1455
      Left            =   120
      TabIndex        =   2
      Top             =   0
      Width           =   4455
   End
   Begin VB.Label Label2 
      Caption         =   "Put this script in the body section of your page."
      Height          =   255
      Left            =   0
      TabIndex        =   11
      Top             =   3960
      Width           =   4215
   End
End
Attribute VB_Name = "frmScript"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdGenerate_Click()
Dim strLang As String
Dim strCode As String
Dim strQ As String

strQ = Label1.Caption
strCode = txtScript.Text
If optVBScript.Value = True Then
strLang = "VBScript"
Else
strLang = "JavaScript"
End If

If optVBScript.Value = True Or optJava.Value = True Then
tag1 = "<p><script language=" & strQ & strLang & strQ & "><!--"
tag2 = strCode
tag3 = "// --></script></p>"
txtScriptCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3
Else
strLang = txtOther.Text
tag1 = "<p><script language=" & strQ & "PHP" & strQ & ">" & strCode & "</script></p>"
txtScriptCode.Text = tag1
End If
Form1.Label13.Caption = "GO"
End Sub

Private Sub Command1_Click()
txtScriptCode.Text = ""
End Sub

Private Sub Command2_Click()
Clipboard.SetText (txtScriptCode.Text)
End Sub
