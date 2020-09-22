VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "HTML Generation Wizard"
   ClientHeight    =   7200
   ClientLeft      =   3600
   ClientTop       =   2670
   ClientWidth     =   10545
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10545
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4320
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdEditCode 
      Caption         =   "Go Edit Code"
      Height          =   375
      Left            =   8040
      TabIndex        =   36
      Top             =   4680
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   375
      Left            =   9360
      TabIndex        =   35
      Top             =   4680
      Width           =   1095
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear"
      Height          =   375
      Left            =   6600
      TabIndex        =   34
      Top             =   4680
      Width           =   1335
   End
   Begin VB.CommandButton cmdGen 
      Caption         =   "&Generate Code"
      Height          =   375
      Left            =   4920
      TabIndex        =   33
      Top             =   4680
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Caption         =   "Links"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   4920
      TabIndex        =   24
      Top             =   5280
      Width           =   5415
      Begin VB.TextBox txtLink3 
         Height          =   285
         Left            =   840
         TabIndex        =   32
         Text            =   "http://www.davis-familysite.com/alex/"
         Top             =   1320
         Width           =   4455
      End
      Begin VB.TextBox txtLink2 
         Height          =   285
         Left            =   840
         TabIndex        =   30
         Text            =   "http://www.davis-familysite.com/alex/"
         Top             =   960
         Width           =   4455
      End
      Begin VB.TextBox txtLink1 
         Height          =   285
         Left            =   840
         TabIndex        =   28
         Text            =   "http://www.davis-familysite.com/alex/"
         Top             =   600
         Width           =   4455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "^"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4680
         TabIndex        =   26
         Top             =   240
         Width           =   375
      End
      Begin VB.Label Label14 
         Caption         =   "Link 3:"
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
         Left            =   120
         TabIndex        =   31
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Link 2:"
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
         Left            =   120
         TabIndex        =   29
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label12 
         Caption         =   "Link 1:"
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
         Left            =   120
         TabIndex        =   27
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label11 
         Caption         =   "Copy and Paste the URL of each link below."
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
         Left            =   240
         TabIndex        =   25
         Top             =   240
         Width           =   4935
      End
   End
   Begin VB.Frame frmBody 
      Caption         =   "&Body Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      TabIndex        =   12
      Top             =   3360
      Width           =   4575
      Begin VB.TextBox txtBodyText 
         Height          =   1455
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   23
         Top             =   2160
         Width           =   4335
      End
      Begin VB.ComboBox cboJustify 
         Height          =   315
         ItemData        =   "frmMain.frx":0000
         Left            =   1200
         List            =   "frmMain.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtFColor 
         Height          =   285
         Left            =   960
         TabIndex        =   19
         Text            =   "Black"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtFSize 
         Height          =   285
         Left            =   960
         TabIndex        =   16
         Text            =   "3"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtFont 
         Height          =   285
         Left            =   600
         TabIndex        =   14
         Text            =   "Arial"
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label10 
         Caption         =   "Body Text: (Type ""<p>"" when you want to make a new line.)"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   22
         Top             =   1680
         Width           =   4335
      End
      Begin VB.Label Label9 
         Caption         =   "Justify Text: "
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
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Font Color:"
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
         Left            =   120
         TabIndex        =   18
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "(1-10)"
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
         Left            =   2400
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Font Size:"
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
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Font"
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
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtHTML 
      Height          =   4455
      Left            =   4920
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      ToolTipText     =   "Your HTML will show up here"
      Top             =   120
      Width           =   5535
   End
   Begin VB.Frame Frame1 
      Caption         =   "&Format Information"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.CheckBox Check1 
         Caption         =   "&Local Image"
         Height          =   255
         Left            =   2400
         TabIndex        =   41
         Top             =   960
         Value           =   1  'Checked
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Height          =   255
         Left            =   3960
         TabIndex        =   40
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox txtBackroundImage 
         Height          =   285
         Left            =   1680
         TabIndex        =   39
         Top             =   600
         Width           =   2175
      End
      Begin VB.TextBox txtALink 
         Height          =   285
         Left            =   1800
         TabIndex        =   10
         Text            =   "Green"
         Top             =   2640
         Width           =   1335
      End
      Begin VB.TextBox txtVLink 
         Height          =   285
         Left            =   1800
         TabIndex        =   8
         Text            =   "Red"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtLink 
         Height          =   285
         Left            =   1800
         TabIndex        =   6
         Text            =   "Blue"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox txtBackColor 
         Height          =   285
         Left            =   1800
         TabIndex        =   4
         Text            =   "White"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox txtTitle 
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Text            =   "New Page"
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label4 
         Caption         =   "Bacround Image:"
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
         Left            =   120
         TabIndex        =   38
         Top             =   600
         Width           =   1815
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   240
         X2              =   4200
         Y1              =   1680
         Y2              =   1680
      End
      Begin VB.Label lblALink 
         Caption         =   "Active  Link Color:"
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
         Left            =   120
         TabIndex        =   9
         Top             =   2640
         Width           =   1695
      End
      Begin VB.Label lblVLink 
         Caption         =   "Visited Link Color:"
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
         Left            =   120
         TabIndex        =   7
         Top             =   2280
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "Link Color:"
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
         Left            =   720
         TabIndex        =   5
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Backround Color:"
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
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "Title:"
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
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Label lblA 
      Caption         =   """"
      Height          =   495
      Left            =   5160
      TabIndex        =   37
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Menu mnupopup 
      Caption         =   "mnupopup"
      Visible         =   0   'False
      Begin VB.Menu mnulinkafter 
         Caption         =   "Links After Text"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnulunkbefore 
         Caption         =   "Links Before Text"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chkBackroundImage_Click()

End Sub

Private Sub cmdClear_Click()
txtHTML.Text = ""
End Sub

Private Sub cmdCopy_Click()
MsgBox "This will copy your code to the clipbord.", vbInformation, "Copy"
Clipboard.SetText (txtHTML.Text)
End Sub




Private Function INtag() As Boolean
If InStrRev(Form1.Text1.Text, "<", Form1.Text1.SelStart, vbTextCompare) > InStrRev(Form1.Text1.Text, ">", Form1.Text1.SelStart, vbTextCompare) Then INtag = True
End Function
Private Function INpropval() As Boolean
Dim x, y As Long
x = InStrRev(Form1.Text1.Text, """", Form1.Text1.SelStart, vbTextCompare)
y = InStrRev(Form1.Text1.Text, "=", Form1.Text1.SelStart, vbTextCompare)
If x > y Then
If InStrRev(Form1.Text1.Text, """", x - 1, vbTextCompare) < InStrRev(text11.Text, "=", x - 1, vbTextCompare) Then INpropval = True
End If
End Function

Private Sub cmdEditCode_Click()

If txtHTML.Text = "" Then
    Dim response
    response = MsgBox("You have not generated any code yet. Do you want" & vbCrLf & "to continue and make your own code?", vbYesNo + vbQuestion, "Continue?")
        If response = vbYes Then
        Unload Me
        Form1.Show
        ElseIf response = vbNo Then
        'do nothing!
        End If
Else
Form1.Text1.Text = txtHTML.Text
Label13.Caption = "GO"
Unload Me
Form1.Show
End If
End Sub

Private Sub cmdExit_Click()
Unload Me
End
End Sub

Private Sub cmdGen_Click()
Dim strlink1 As String
Dim strlink2 As String
Dim strlink3 As String
Dim strTitle As String
Dim strBGColor As String
Dim strBGImage As String
Dim strLink As String
Dim strVlink As String
Dim strAlink As String
Dim strfontname As String
Dim strfontcolor As String
Dim strfontsize As String
Dim strtext As String
Dim strQ As String
strQ = lblA.Caption

strlink1 = txtLink1.Text
strlink2 = txtLink2.Text
strlink3 = txtLink3.Text
strTitle = txtTitle.Text
strBGColor = txtBackColor.Text
strLink = txtLink.Text
strVlink = txtVLink.Text
strAlink = txtALink.Text
strfontname = txtFont.Text
strfontsize = txtFSize.Text
strfontcolor = txtFColor.Text
strtext = txtBodyText.Text


If txtBackroundImage.Text > "" Then
tag1 = "<html>"
tag2 = "<head>"
tag3 = "<title>" & strTitle & "</title>"
tag4 = "</head>"
tag5 = "<body background = " & strQ & txtBackroundImage.Text & strQ & " link = " & strQ & strLink & strQ & " vlink = " & strQ & strVlink & strQ & " alink =" & strQ & strAlink & strQ & " >"
tag6 = "<font face = " & strQ & strfontname & strQ & " size = " & strQ & strfontsize & strQ & " color = " & strQ & strfontcolor & strQ & ">"
tagA = "<" & cboJustify.Text & ">"
tag7 = txtBodyText.Text & "<p>"
tagb = "</" & cboJustify.Text & ">"
tag13 = "</font>"
tag8 = "<a href = " & strQ & strlink1 & strQ & "> Link 1 Title Goes Here </a> <p>"
tag9 = "<a href = " & strQ & strlink2 & strQ & "> Link 2 Title Goes Here </a> <p>"
tag10 = "<a href = " & strQ & strlink3 & strQ & "> Link 3 Title Goes Here </a> <p>"
tag11 = "</body>"
tag12 = "</html>"

If mnulinkafter.Checked = True Then
txtHTML.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tagA & tag7 & tagb & vbCrLf & tag13 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12

Else
txtHTML.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tagA & tag7 & tagb & vbCrLf & tag13 & vbCrLf & vbCrLf & tag11 & vbCrLf & tag12
End If
Else
tag1 = "<html>"
tag2 = "<head>"
tag3 = "<title>" & strTitle & "</title>"
tag4 = "</head>"
tag5 = "<body bgcolor = " & strQ & strBGColor & strQ & " link = " & strQ & strLink & strQ & " vlink = " & strQ & strVlink & strQ & " alink =" & strQ & strAlink & strQ & " >"
tag6 = "<font face = " & strQ & strfontname & strQ & " size = " & strQ & strfontsize & strQ & " color = " & strQ & strfontcolor & strQ & ">"
tagA = "<" & cboJustify.Text & ">"
tag7 = txtBodyText.Text & "<p>"
tagb = "</" & cboJustify.Text & ">"
tag13 = "</font>"
tag8 = "<a href = " & strQ & strlink1 & strQ & "> Link 1 Title Goes Here </a> <p>"
tag9 = "<a href = " & strQ & strlink2 & strQ & "> Link 2 Title Goes Here </a> <p>"
tag10 = "<a href = " & strQ & strlink3 & strQ & "> Link 3 Title Goes Here </a> <p>"
tag11 = "</body>"
tag12 = "</html>"

If mnulinkafter.Checked = True Then
txtHTML.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tagA & tag7 & tagb & vbCrLf & tag13 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12

Else
txtHTML.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tagA & tag7 & tagb & vbCrLf & tag13 & vbCrLf & vbCrLf & tag11 & vbCrLf & tag12
End If
End If
End Sub

Private Sub Command1_Click()
PopupMenu mnupopup
End Sub

Private Sub Command2_Click()
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
txtBackroundImage.Text = sFile
End Sub

Private Sub mnulinkafter_Click()
mnulinkafter.Checked = True
mnulunkbefore.Checked = False
txtLink1.SetFocus
End Sub

Private Sub mnulunkbefore_Click()
mnulunkbefore.Checked = True
mnulinkafter.Checked = False
txtLink1.SetFocus
End Sub

