VERSION 5.00
Begin VB.Form frmBasic 
   Caption         =   "Basic HTML Tags"
   ClientHeight    =   3345
   ClientLeft      =   6120
   ClientTop       =   4095
   ClientWidth     =   6060
   Icon            =   "frmBasic.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3345
   ScaleWidth      =   6060
   Begin VB.CommandButton cmdcopy 
      Caption         =   "C&opy"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtCode 
      Height          =   1695
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   1
      Text            =   "frmBasic.frx":030A
      Top             =   720
      Width           =   5775
   End
   Begin VB.ComboBox cboFunction 
      Height          =   315
      ItemData        =   "frmBasic.frx":0337
      Left            =   2160
      List            =   "frmBasic.frx":0347
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   """"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   1320
      Width           =   1335
   End
   Begin VB.Label lblHint 
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   2535
   End
End
Attribute VB_Name = "frmBasic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboFunction_Click()
Dim strQ As String
strQ = Label1.Caption

If cboFunction.Text = "Select an Item..." Then
    txtCode.Text = "Click on items on the list to view the code."
    lblHint.Caption = ""
ElseIf cboFunction.Text = "Bullet List" Then
    txtCode.Text = "<ul>" & vbCrLf & "<li>Bullet One</li>" & vbCrLf & "<li>Bullet Two</li>" & vbCrLf & "<li>Bullet Three</li>" & vbCrLf & "</ul>"
    lblHint.Caption = ""
ElseIf cboFunction.Text = "Number List" Then
    txtCode.Text = "<ol>" & vbCrLf & "<li>Number One</li>" & vbCrLf & "<li>Number Two</li>" & vbCrLf & "<li>Number Three</li>" & vbCrLf & "</ol>"
    lblHint.Caption = ""
ElseIf cboFunction.Text = "Scrolling Marqee" Then
    txtCode.Text = "<p><marquee bgcolor=" & strQ & "YOUR BACK COLOR" & strQ & " border=" & strQ & "0" & strQ & ">YOUR TEXT</marquee></p>"
Else
End If
End Sub

Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdCopy_Click()
Clipboard.SetText (txtCode.Text)
End Sub

Private Sub Form_Load()
cboFunction.ListIndex = 0
End Sub
