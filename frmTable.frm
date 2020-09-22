VERSION 5.00
Begin VB.Form frmTable 
   Caption         =   "Insert Table"
   ClientHeight    =   4530
   ClientLeft      =   6120
   ClientTop       =   3390
   ClientWidth     =   6075
   Icon            =   "frmTable.frx":0000
   LinkTopic       =   "Form4"
   ScaleHeight     =   4530
   ScaleWidth      =   6075
   Begin VB.CommandButton Command1 
      Caption         =   "Copy Text"
      Height          =   375
      Left            =   1680
      TabIndex        =   21
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "&Generate"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   375
      Left            =   4560
      TabIndex        =   14
      Top             =   4080
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Code"
      Height          =   2055
      Left            =   120
      TabIndex        =   11
      Top             =   1920
      Width           =   5775
      Begin VB.TextBox txtCode 
         Height          =   1455
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   3  'Both
         TabIndex        =   13
         Top             =   480
         Width           =   5415
      End
      Begin VB.Label Label8 
         Caption         =   """"
         Height          =   735
         Left            =   2160
         TabIndex        =   18
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label6 
         Caption         =   "Place this code where you want the table to appear."
         Height          =   255
         Left            =   840
         TabIndex        =   12
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Table Information"
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5775
      Begin VB.TextBox txtTableWidth 
         Height          =   285
         Left            =   2520
         TabIndex        =   20
         Text            =   "100"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtColor 
         Height          =   285
         Left            =   4320
         TabIndex        =   17
         Text            =   "Silver"
         Top             =   960
         Width           =   1335
      End
      Begin VB.ComboBox cboColumn 
         Height          =   315
         ItemData        =   "frmTable.frx":030A
         Left            =   4320
         List            =   "frmTable.frx":031D
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cboRows 
         Height          =   315
         ItemData        =   "frmTable.frx":0330
         Left            =   4320
         List            =   "frmTable.frx":0343
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtPad 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Text            =   "1"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtSpacing 
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Text            =   "0"
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtBorder 
         Height          =   285
         Left            =   1080
         TabIndex        =   4
         Text            =   "1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label9 
         Caption         =   "Table Width: (%)"
         Height          =   255
         Left            =   1200
         TabIndex        =   19
         Top             =   1320
         Width           =   1695
      End
      Begin VB.Label Label7 
         Caption         =   "Border Color"
         Height          =   255
         Left            =   2640
         TabIndex        =   16
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         Caption         =   "Number of Columns"
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label4 
         Caption         =   "Number of Rows:"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label3 
         Caption         =   "Cell Padding:"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label2 
         Caption         =   "Cell Spacing:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Border width:"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
End Sub

Private Sub cmdGenerate_Click()
Dim strwidth As String
Dim strpad As String
Dim strendrow As String
Dim strspace As String
Dim strColor As String
Dim strRow As String
Dim strCol As String
Dim strTW As String

Dim strQ As String

strQ = Label8.Caption
strwidth = txtBorder.Text
strpad = txtPad.Text
strspace = txtSpacing.Text
strColor = txtColor.Text
strRow = "<tr>"
strCol = "<td> Your Text Here </td>"
strendrow = "</tr>"
strTW = txtTableWidth.Text
tagb = "<table border=" & strQ & strwidth & strQ & " cellpadding=" & strQ & strpad & strQ & " cellspacing=" & strQ & strspace & strQ & " width=" & strQ & strTW & "%" & strQ & " bordercolor =" & strQ & strColor & strQ & ">"

If cboRows.Text = "1" And cboColumn.Text = "1" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strendrow
tag5 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5
ElseIf cboRows.Text = "1" And cboColumn.Text = "2" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strendrow
tag6 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6
ElseIf cboRows.Text = "1" And cboColumn.Text = "3" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strendrow
tag7 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7
ElseIf cboRows.Text = "1" And cboColumn.Text = "4" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strCol
tag7 = strendrow
tag8 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8
ElseIf cboRows.Text = "1" And cboColumn.Text = "5" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strCol
tag7 = strCol
tag8 = strendrow
tag9 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9
ElseIf cboRows.Text = "2" And cboColumn.Text = "1" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strendrow
tag5 = strRow
tag6 = strCol
tag7 = strendrow
tag8 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8
ElseIf cboRows.Text = "2" And cboColumn.Text = "2" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strendrow
tag6 = strRow
tag7 = strCol
tag8 = strCol
tag9 = strendrow
tag10 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10
ElseIf cboRows.Text = "2" And cboColumn.Text = "3" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strendrow
tag7 = strRow
tag8 = strCol
tag9 = strCol
tag10 = strCol
tag11 = strendrow
tag12 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12
ElseIf cboRows.Text = "2" And cboColumn.Text = "4" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strCol
tag7 = strendrow
tag8 = strRow
tag9 = strCol
tag10 = strCol
tag11 = strCol
tag12 = strCol
tag13 = strendrow
tag14 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbCrLf & tag13 & vbCrLf & tag14
ElseIf cboRows.Text = "2" And cboColumn.Text = "5" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strCol
tag7 = strCol
tag8 = strendrow
tag9 = strRow
tag10 = strCol
tag11 = strCol
tag12 = strCol
tag13 = strCol
tag14 = strCol
tag15 = strendrow
tag16 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbCrLf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16
ElseIf cboRows.Text = "3" And cboColumn.Text = "1" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strendrow
tag5 = strRow
tag6 = strCol
tag7 = strendrow
tag8 = strRow
tag9 = strCol
tag10 = strenrow
tag11 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11
ElseIf cboRows.Text = "3" And cboColumn.Text = "2" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strendrow
tag6 = strRow
tag7 = strCol
tag8 = strCol
tag9 = strendrow
tag10 = strRow
tag11 = strCol
tag12 = strCol
tag13 = strenrow
tag14 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbCrLf & tag13 & vbCrLf & tag14
ElseIf cboRows.Text = "3" And cboColumn.Text = "3" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strendrow
tag7 = strRow
tag8 = strCol
tag9 = strCol
tag10 = strCol
tag11 = strendrow
tag12 = strRow
tag13 = strCol
tag14 = strCol
tag15 = strCol
tag16 = strenrow
tag17 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbCrLf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16 & vbCrLf & tag17
ElseIf cboRows.Text = "3" And cboColumn.Text = "4" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strCol
tag7 = strendrow
tag8 = strRow
tag9 = strCol
tag10 = strCol
tag11 = strCol
tag12 = strCol
tag13 = strendrow
tag14 = strRow
tag15 = strCol
tag16 = strCol
tag17 = strCol
tag18 = strCol
tag19 = strenrow
tag20 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbCrLf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16 & vbCrLf & tag17 & vbCrLf & tag18 & vbCrLf & tag19 & vbCrLf & tag20
ElseIf cboRows.Text = "3" And cboColumn.Text = "5" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strCol
tag7 = strCol
tag8 = strendrow
tag9 = strRow
tag10 = strCol
tag11 = strCol
tag12 = strCol
tag13 = strCol
tag14 = strCol
tag15 = strendrow
tag16 = strRow
tag17 = strCol
tag18 = strCol
tag19 = strCol
tag20 = strCol
tag21 = strCol
tag22 = strenrow
tag23 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbCrLf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16 & vbCrLf & tag17 & vbCrLf & tag18 & vbCrLf & tag19 & vbCrLf & tag20 & vbCrLf & tag21 & vbCrLf & tag22 & vbCrLf & tag23
ElseIf cboRows.Text = "4" And cboColumn.Text = "1" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strendrow
tag5 = strRow
tag6 = strCol
tag7 = strendrow
tag8 = strRow
tag9 = strCol
tag10 = strenrow
tag11 = strRow
tag12 = strCol
tag13 = strendrow
tag14 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbrclf & tag13 & vbCrLf & tag14
ElseIf cboRows.Text = "4" And cboColumn.Text = "2" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strendrow
tag6 = strRow
tag7 = strCol
tag8 = strCol
tag9 = strendrow
tag10 = strRow
tag11 = strCol
tag12 = strCol
tag13 = strenrow
tag14 = strRow
tag15 = strCol
tag16 = strCol
tag17 = strendrow
tag18 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbrclf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16 & vbCrLf & tag17 & vbCrLf & tag18
ElseIf cboRows.Text = "4" And cboColumn.Text = "3" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strendrow
tag7 = strRow
tag8 = strCol
tag9 = strCol
tag10 = strCol
tag11 = strendrow
tag12 = strRow
tag13 = strCol
tag14 = strCol
tag15 = strCol
tag16 = strenrow
tag17 = strRow
tag18 = strCol
tag19 = strCol
tag20 = strCol
tag21 = strendrow
tag22 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbrclf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16 & vbCrLf & tag17 & vbCrLf & tag18 & vbCrLf & tag19 & vbCrLf & tag20 & vbCrLf & tag21 & vbCrLf & tag22
ElseIf cboRows.Text = "4" And cboColumn.Text = "4" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strCol
tag7 = strendrow
tag8 = strRow
tag9 = strCol
tag10 = strCol
tag11 = strCol
tag12 = strCol
tag13 = strendrow
tag14 = strRow
tag15 = strCol
tag16 = strCol
tag17 = strCol
tag18 = strCol
tag19 = strenrow
tag20 = strRow
tag21 = strCol
tag22 = strCol
tag23 = strCol
tag24 = strCol
tag25 = strendrow
tag26 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbrclf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16 & vbCrLf & tag17 & vbCrLf & tag18 & vbCrLf & tag19 & vbCrLf & tag20 & vbCrLf & tag21 & vbCrLf & tag22 & vbCrLf & tag23 & vbCrLf & tag24 & vbCrLf & tag25 & vbCrLf & tag26
ElseIf cboRows.Text = "4" And cboColumn.Text = "5" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strCol
tag7 = strCol
tag8 = strendrow
tag9 = strRow
tag10 = strCol
tag11 = strCol
tag12 = strCol
tag13 = strCol
tag14 = strCol
tag15 = strendrow
tag16 = strRow
tag17 = strCol
tag18 = strCol
tag19 = strCol
tag20 = strCol
tag21 = strCol
tag22 = strenrow
tag23 = strRow
tag24 = strCol
tag25 = strCol
tag26 = strCol
tag27 = strCol
tag28 = strCol
tag29 = strendrow
tag30 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbrclf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16 & vbCrLf & tag17 & vbCrLf & tag18 & vbCrLf & tag19 & vbCrLf & tag20 & vbCrLf & tag21 & vbCrLf & tag22 & vbCrLf & tag23 & vbCrLf & tag24 & vbCrLf & tag25 & vbCrLf & tag26 & vbCrLf & tag27 & vbCrLf & tag28 & vbCrLf & tag29 & vbCrLf & tag30
ElseIf cboRows.Text = "5" And cboColumn.Text = "1" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strendrow
tag5 = strRow
tag6 = strCol
tag7 = strendrow
tag8 = strRow
tag9 = strCol
tag10 = strendrow
tag11 = strRow
tag12 = strCol
tag13 = strendrow
tag14 = strRow
tag15 = strCol
tag16 = strendrow
tag17 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbrclf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16 & vbCrLf & tag17
ElseIf cboRows.Text = "5" And cboColumn.Text = "2" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strendrow
tag6 = strRow
tag7 = strCol
tag8 = strCol
tag9 = strendrow
tag10 = strRow
tag11 = strCol
tag12 = strCol
tag13 = strendrow
tag14 = strRow
tag15 = strCol
tag16 = strCol
tag17 = strendrow
tag18 = strRow
tag19 = strCol
tag20 = strCol
tag21 = strendrow
tag22 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbrclf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16 & vbCrLf & tag17 & vbCrLf & tag18 & vbCrLf & tag19 & vbCrLf & tag20 & vbCrLf & tag21 & vbCrLf & tag22
ElseIf cboRows.Text = "5" And cboColumn.Text = "3" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strendrow
tag7 = strRow
tag8 = strCol
tag9 = strCol
tag10 = strCol
tag11 = strendrow
tag12 = strRow
tag13 = strCol
tag14 = strCol
tag15 = strCol
tag16 = strendrow
tag17 = strRow
tag18 = strCol
tag19 = strCol
tag20 = strCol
tag21 = strendrow
tag22 = strRow
tag23 = strCol
tag24 = strCol
tag25 = strCol
tag26 = strendrow
tag27 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbrclf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16 & vbCrLf & tag17 & vbCrLf & tag18 & vbCrLf & tag19 & vbCrLf & tag20 & vbCrLf & tag21 & vbCrLf & tag22 & vbCrLf & tag23 & vbCrLf & tag24 & vbCrLf & tag25 & vbCrLf & tag26 & vbCrLf & tag27
ElseIf cboRows.Text = "5" And cboColumn.Text = "4" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strCol
tag7 = strendrow
tag8 = strRow
tag9 = strCol
tag10 = strCol
tag11 = strCol
tag12 = strCol
tag13 = strendrow
tag14 = strRow
tag15 = strCol
tag16 = strCol
tag17 = strCol
tag18 = strCol
tag19 = strendrow
tag20 = strRow
tag21 = strCol
tag22 = strCol
tag23 = strCol
tag24 = strCol
tag25 = strendrow
tag26 = strRow
tag27 = strCol
tag28 = strCol
tag29 = strCol
tag30 = strCol
tag31 = strendrow
tag32 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbrclf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16 & vbCrLf & tag17 & vbCrLf & tag18 & vbCrLf & tag19 & vbCrLf & tag20 & vbCrLf & tag21 & vbCrLf & tag22 & vbCrLf & tag23 & vbCrLf & tag24 & vbCrLf & tag25 & vbCrLf & tag26 & vbCrLf & tag27 & vbCrLf & tag28 & vbCrLf & tag29 & vbCrLf & tag30 & vbCrLf & tag31 & vbCrLf & tag32
ElseIf cboRows.Text = "5" And cboColumn.Text = "4" Then
tag1 = tagb
tag2 = strRow
tag3 = strCol
tag4 = strCol
tag5 = strCol
tag6 = strCol
tag7 = strCol
tag8 = strendrow
tag9 = strRow
tag10 = strCol
tag11 = strCol
tag12 = strCol
tag13 = strCol
tag14 = strCol
tag15 = strendrow
tag16 = strRow
tag17 = strCol
tag18 = strCol
tag19 = strCol
tag20 = strCol
tag21 = strCol
tag22 = strendrow
tag23 = strRow
tag24 = strCol
tag25 = strCol
tag26 = strCol
tag27 = strCol
tag28 = strCol
tag29 = strendrow
tag30 = strRow
tag31 = strCol
tag32 = strCol
tag33 = strCol
tag34 = strCol
tag35 = strCol
tag36 = strendrow
tag37 = "</table>"
txtCode.Text = tag1 & vbCrLf & tag2 & vbCrLf & tag3 & vbCrLf & tag4 & vbCrLf & tag5 & vbCrLf & tag6 & vbCrLf & tag7 & vbCrLf & tag8 & vbCrLf & tag9 & vbCrLf & tag10 & vbCrLf & tag11 & vbCrLf & tag12 & vbrclf & tag13 & vbCrLf & tag14 & vbCrLf & tag15 & vbCrLf & tag16 & vbCrLf & tag17 & vbCrLf & tag18 & vbCrLf & tag19 & vbCrLf & tag20 & vbCrLf & tag21 & vbCrLf & tag22 & vbCrLf & tag23 & vbCrLf & tag24 & vbCrLf & tag25 & vbCrLf & tag26 & vbCrLf & tag27 & vbCrLf & tag28 & vbCrLf & tag29 & vbCrLf & tag30 & vbCrLf & tag31 & vbCrLf & tag32 & vbCrLf & tag33 & vbCrLf & tag34 & vbCrLf & tag35 & vbCrLf & tag36 & vbCrLf & tag37
Else
MsgBox "Error Developing Code", vbCritical, "Error"
End If
Form1.Label13.Caption = "GO"
End Sub

Private Sub Command1_Click()
Clipboard.SetText (txtCode.Text)
End Sub

Private Sub Form_Load()
cboRows.ListIndex = 0
cboColumn.ListIndex = 0
End Sub

Private Sub txtTableWidth_Change()
If txtTableWidth.Text = "" And txtTableWidth.Text >= 1 And txtTableWidth.Text <= 100 Then
Else
    If IsNumeric(txtTableWidth.Text) = False Then
    MsgBox "Please write a number between 1 and 100", vbCritical, "error"
    Else
    End If
End If
End Sub
