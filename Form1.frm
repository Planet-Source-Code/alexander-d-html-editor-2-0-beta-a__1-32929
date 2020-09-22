VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "HTML Editor"
   ClientHeight    =   7575
   ClientLeft      =   3195
   ClientTop       =   2250
   ClientWidth     =   12300
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7575
   ScaleWidth      =   12300
   Begin VB.PictureBox Picture5 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8040
      ScaleHeight     =   315
      ScaleWidth      =   3915
      TabIndex        =   31
      Top             =   0
      Width           =   3975
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "0 Tags"
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
         TabIndex        =   33
         ToolTipText     =   "Click syntax to update tag count"
         Top             =   45
         Width           =   1335
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "0 Characters"
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
         Left            =   1800
         TabIndex        =   32
         Top             =   45
         Width           =   2055
      End
   End
   Begin VB.Timer Timer2 
      Interval        =   25
      Left            =   4560
      Top             =   3000
   End
   Begin VB.PictureBox picLines 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   2520
      ScaleHeight     =   1455
      ScaleWidth      =   615
      TabIndex        =   30
      Top             =   480
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Syntax"
      Height          =   495
      Left            =   3000
      TabIndex        =   26
      Top             =   2040
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox Text1 
      Height          =   1455
      Left            =   3120
      TabIndex        =   21
      Top             =   480
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   2566
      _Version        =   393217
      BackColor       =   16777215
      ScrollBars      =   3
      TextRTF         =   $"Form1.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture3 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   2640
      ScaleHeight     =   495
      ScaleWidth      =   14535
      TabIndex        =   13
      Top             =   0
      Width           =   14535
      Begin MSComctlLib.Toolbar tbToolBar 
         Height          =   390
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   9135
         _ExtentX        =   16113
         _ExtentY        =   688
         ButtonWidth     =   609
         ButtonHeight    =   582
         Appearance      =   1
         ImageList       =   "imlToolbarIcons"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   19
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "New"
               Object.ToolTipText     =   "New"
               ImageKey        =   "New"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Open"
               Object.ToolTipText     =   "Open"
               ImageKey        =   "Open"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Save"
               Object.ToolTipText     =   "Save"
               ImageKey        =   "Save"
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Print"
               Object.ToolTipText     =   "Print"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Cut"
               Object.ToolTipText     =   "Cut"
               ImageKey        =   "Cut"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Copy"
               Object.ToolTipText     =   "Copy"
               ImageKey        =   "Copy"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Paste"
               Object.ToolTipText     =   "Paste"
               ImageKey        =   "Paste"
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Delete"
               Object.ToolTipText     =   "Delete"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Undo"
               Object.ToolTipText     =   "Undo"
               ImageKey        =   "Undo"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Redo"
               Object.ToolTipText     =   "Redo"
               ImageKey        =   "Redo"
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Find"
               Object.ToolTipText     =   "Find"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Tab Right"
               Object.ToolTipText     =   "File Format"
               ImageKey        =   "Tab Right"
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Text Files (.txt)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Help"
               Object.ToolTipText     =   "Help"
               ImageKey        =   "Help"
            EndProperty
         EndProperty
      End
   End
   Begin VB.Timer Timer1 
      Left            =   4200
      Top             =   3000
   End
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   7320
      Width           =   12300
      _ExtentX        =   21696
      _ExtentY        =   450
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   7
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            AutoSize        =   1
            Bevel           =   0
            Object.Width           =   13529
            MinWidth        =   3951
            Text            =   "Welcome to the HTML editor"
            TextSave        =   "Welcome to the HTML editor"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   1
            Enabled         =   0   'False
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "CAPS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   2
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "NUM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   3
            Enabled         =   0   'False
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "INS"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   4
            Enabled         =   0   'False
            Object.Width           =   1129
            MinWidth        =   1129
            TextSave        =   "SCRL"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel6 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   5
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "4:29 PM"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel7 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Style           =   6
            Object.Width           =   1482
            MinWidth        =   1482
            TextSave        =   "3/22/02"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open HTML file"
      FileName        =   "Page One"
      Filter          =   "Hypertext Files (*.html)|*.html|All Files (*.*)|*.*|Text Files (*.txt*)|*.txt*|Hypertext Files (*.htm)|*.htm*"
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   11655
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   20558
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   661
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Browser"
      TabPicture(0)   =   "Form1.frx":03EA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Dir1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "File1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Drive1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "<  >"
      TabPicture(1)   =   "Form1.frx":0406
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "tv1"
      Tab(1).Control(1)=   "Label4"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Help"
      TabPicture(2)   =   "Form1.frx":0422
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Picture4"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Func."
      TabPicture(3)   =   "Form1.frx":043E
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Picture2"
      Tab(3).Control(1)=   "Combo2"
      Tab(3).ControlCount=   2
      Begin VB.PictureBox Picture2 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   -74880
         ScaleHeight     =   375
         ScaleWidth      =   2295
         TabIndex        =   28
         Top             =   480
         Width           =   2295
         Begin VB.Label Label6 
            Caption         =   "Functions:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.ComboBox Combo2 
         BackColor       =   &H00DCF8F5&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A26C1E&
         Height          =   2580
         ItemData        =   "Form1.frx":045A
         Left            =   -74880
         List            =   "Form1.frx":046A
         Style           =   1  'Simple Combo
         TabIndex        =   27
         Text            =   "Combo2"
         Top             =   480
         Width           =   2295
      End
      Begin ComctlLib.TreeView tv1 
         Height          =   3975
         Left            =   -74880
         TabIndex        =   25
         Top             =   1080
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   7011
         _Version        =   327682
         LabelEdit       =   1
         Style           =   7
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         Height          =   11055
         Left            =   -74880
         ScaleHeight     =   10995
         ScaleWidth      =   2235
         TabIndex        =   14
         Top             =   480
         Width           =   2295
         Begin VB.Line Line2 
            X1              =   240
            X2              =   1920
            Y1              =   3480
            Y2              =   3480
         End
         Begin VB.Label Label12 
            BackStyle       =   0  'Transparent
            Caption         =   """?"" Means arbitrary number. "":::"" means arbitrary date."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   120
            TabIndex        =   20
            Top             =   3840
            Width           =   1935
         End
         Begin VB.Label Label11 
            BackColor       =   &H00FFFFFF&
            Caption         =   "NOTE IN TAGS:"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   3600
            Width           =   2055
         End
         Begin VB.Label Label10 
            BackStyle       =   0  'Transparent
            Caption         =   $"Form1.frx":0499
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
            Left            =   120
            TabIndex        =   18
            Top             =   2040
            Width           =   2055
         End
         Begin VB.Label Label9 
            BackStyle       =   0  'Transparent
            Caption         =   "Q:  Why can't I see my HTML files in the file browser?"
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
            TabIndex        =   17
            Top             =   1320
            Width           =   2055
         End
         Begin VB.Label Label8 
            BackColor       =   &H00FFFFFF&
            Caption         =   "F.A.Q."
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00558E2F&
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   960
            Width           =   1935
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Help"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   27.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   2175
         End
      End
      Begin VB.DriveListBox Drive1 
         BackColor       =   &H00DCF8F5&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A26C1E&
         Height          =   330
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   2295
      End
      Begin VB.FileListBox File1 
         BackColor       =   &H00DCF8F5&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A26C1E&
         Height          =   1980
         Left            =   120
         Pattern         =   "*.html*"
         TabIndex        =   10
         Top             =   2880
         Width           =   2295
      End
      Begin VB.DirListBox Dir1 
         BackColor       =   &H00DCF8F5&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00A26C1E&
         Height          =   1890
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label4 
         Caption         =   "Place your Cursor where you want the tags to appear."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   -74880
         TabIndex        =   12
         Top             =   480
         Width           =   2295
      End
   End
   Begin VB.TextBox Text2 
      Height          =   735
      Left            =   240
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   3720
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   855
      Left            =   360
      TabIndex        =   6
      Top             =   3360
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   3000
      TabIndex        =   22
      Top             =   840
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   18
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0534
            Key             =   ""
            Object.Tag             =   "&New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0690
            Key             =   ""
            Object.Tag             =   "Insert I&mage"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0700
            Key             =   ""
            Object.Tag             =   "N&ew Frames Page"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":085C
            Key             =   ""
            Object.Tag             =   "Your Page"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":09B8
            Key             =   ""
            Object.Tag             =   "&Help..."
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0B14
            Key             =   ""
            Object.Tag             =   "&Open"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0C70
            Key             =   ""
            Object.Tag             =   "Save &As"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0DCC
            Key             =   ""
            Object.Tag             =   "&Copy"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0F28
            Key             =   ""
            Object.Tag             =   "C&ut"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1084
            Key             =   ""
            Object.Tag             =   "&Paste"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":11E0
            Key             =   ""
            Object.Tag             =   "&Delete"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":133C
            Key             =   ""
            Object.Tag             =   "Syntax All"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1498
            Key             =   ""
            Object.Tag             =   "&Style Editor"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18C8
            Key             =   ""
            Object.Tag             =   "S&cript Editor"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1CF8
            Key             =   ""
            Object.Tag             =   "Insert &Table"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D68
            Key             =   ""
            Object.Tag             =   "&Horizontal Rule"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DD8
            Key             =   ""
            Object.Tag             =   "Our &Website"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1ED8
            Key             =   ""
            Object.Tag             =   "Look for *.htm* files"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlToolbarIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2274
            Key             =   "New"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2386
            Key             =   "Open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2498
            Key             =   "Save"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":25AA
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26BC
            Key             =   "Cut"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":27CE
            Key             =   "Copy"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":28E0
            Key             =   "Paste"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29F2
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2B04
            Key             =   "Undo"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C16
            Key             =   "Redo"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2D28
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2E3A
            Key             =   "Tab Left"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2F4C
            Key             =   "Tab Right"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":305E
            Key             =   "Help"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label14 
      Caption         =   "1"
      Height          =   615
      Left            =   2880
      TabIndex        =   24
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "GO"
      Height          =   255
      Left            =   3480
      TabIndex        =   23
      Top             =   1200
      Width           =   1095
   End
   Begin ComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   3
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3170
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3882
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":3F34
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList imlIcons2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4286
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":45D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":492A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4C7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":4FCE
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5320
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5672
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":59C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":5D16
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6068
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":63BA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":670C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6A5E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":6DB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7102
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7454
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":77A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7AF8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":7E4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":819C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":84EE
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form1.frx":8840
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   615
      Left            =   3000
      TabIndex        =   5
      Top             =   840
      Width           =   1095
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Files that can be loaded:"
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
      Left            =   0
      TabIndex        =   2
      Top             =   3000
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   3120
      TabIndex        =   0
      Top             =   840
      Width           =   1095
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu o 
         Caption         =   "&Open"
      End
      Begin VB.Menu n 
         Caption         =   "&New"
      End
      Begin VB.Menu s 
         Caption         =   "&Save"
         Visible         =   0   'False
      End
      Begin VB.Menu sa 
         Caption         =   "Save &As"
      End
      Begin VB.Menu close 
         Caption         =   "&Close"
      End
      Begin VB.Menu pcode 
         Caption         =   "&Print"
      End
      Begin VB.Menu vbdiv11 
         Caption         =   "-"
      End
      Begin VB.Menu import 
         Caption         =   "Import &HTML"
         Enabled         =   0   'False
      End
      Begin VB.Menu vbdiv2 
         Caption         =   "-"
      End
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu edit 
      Caption         =   "&Edit"
      Begin VB.Menu udo 
         Caption         =   "Undo"
      End
      Begin VB.Menu rdo 
         Caption         =   "Redo"
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu cut 
         Caption         =   "C&ut"
         Shortcut        =   ^X
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu insrt 
         Caption         =   "Insert"
         Visible         =   0   'False
         Begin VB.Menu date 
            Caption         =   "Date"
         End
         Begin VB.Menu time 
            Caption         =   "Time"
         End
      End
      Begin VB.Menu vbdiv 
         Caption         =   "-"
      End
      Begin VB.Menu selall 
         Caption         =   "Select All"
      End
      Begin VB.Menu del 
         Caption         =   "&Delete"
      End
   End
   Begin VB.Menu view 
      Caption         =   "&View"
      Begin VB.Menu yp 
         Caption         =   "Your Page"
         Shortcut        =   {F4}
      End
      Begin VB.Menu stbar 
         Caption         =   "Status Bar"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
      Begin VB.Menu fbdick 
         Caption         =   "File Browser"
         Checked         =   -1  'True
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnu 
      Caption         =   "popup"
      Visible         =   0   'False
      Begin VB.Menu delete 
         Caption         =   "&Delete"
      End
      Begin VB.Menu rname 
         Caption         =   "&Rename"
      End
      Begin VB.Menu rf 
         Caption         =   "Refr&esh"
      End
   End
   Begin VB.Menu mnuextras 
      Caption         =   "&Tools"
      Begin VB.Menu mnuframse 
         Caption         =   "N&ew Frames Page"
      End
      Begin VB.Menu vbdiv1 
         Caption         =   "-"
      End
      Begin VB.Menu mnustyle 
         Caption         =   "&Style Editor"
      End
      Begin VB.Menu mnuscript 
         Caption         =   "S&cript Editor"
      End
      Begin VB.Menu mnuimage 
         Caption         =   "Insert I&mage"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnutable 
         Caption         =   "Insert &Table"
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu HR 
         Caption         =   "&Horizontal Rule"
      End
      Begin VB.Menu mnubasicfunctions 
         Caption         =   "Basic Tags..."
         Shortcut        =   ^T
      End
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "mnupopup"
      Visible         =   0   'False
      Begin VB.Menu htm 
         Caption         =   "Look for *.htm* files"
      End
      Begin VB.Menu html 
         Caption         =   "Look for *.html* files"
         Checked         =   -1  'True
      End
      Begin VB.Menu txtfiles 
         Caption         =   "Look For *.txt* files"
      End
      Begin VB.Menu divide1 
         Caption         =   "-"
      End
      Begin VB.Menu allf 
         Caption         =   "Look for ALL files"
      End
   End
   Begin VB.Menu sytax1 
      Caption         =   "&Syntax"
      Begin VB.Menu sytaxall 
         Caption         =   "Syntax All"
         Shortcut        =   ^S
      End
      Begin VB.Menu sytaxnone 
         Caption         =   "Syntax None"
         Checked         =   -1  'True
         Shortcut        =   ^U
      End
   End
   Begin VB.Menu mnuhelp 
      Caption         =   "&Help"
      Begin VB.Menu aboutus 
         Caption         =   "A&bout"
      End
      Begin VB.Menu divider 
         Caption         =   "-"
      End
      Begin VB.Menu mnuhelphelp 
         Caption         =   "&Help..."
      End
      Begin VB.Menu mnusite 
         Caption         =   "Our &Website"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim UndoStack() As String, UndoStage, Undoing
Dim apppath As String
Dim starttime As Date
Dim tmpchr As String * 1
Dim tmpint As Long
Dim Color(3) As Variant
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Long, lparam As Any) As Long
Private Const EM_GETLINECOUNT = &HBA
Private Const EM_LINEINDEX = &HBB
Private Const EM_LINELENGTH = &HC1
Private Const EM_GETFIRSTVISIBLELINE = &HCE





Private Sub Combo1_DblClick()
Dim strString As String
strString = Combo1.Text
Text1.SelText = strString

End Sub

Private Sub allf_Click()
File1.Pattern = "*.*"
End Sub

Private Sub close_Click()
Dim response
response = MsgBox("By doing this you will loose all unsaved work. Save work now?", vbYesNoCancel + vbQuestion, "Save?")

If response = vbYes Then
sa_Click
ElseIf response = vbNo Then
Text1.Text = ""
Else
End If
End Sub

Private Sub Combo2_DblClick()
If Combo2.Text = "(New Line)" Then
Text1.SelText = vbCrLf
ElseIf Combo2.Text = "(Insert Code)" Then
On Error Resume Next
Text1.SelText = Clipboard.GetText
ElseIf Combo2.Text = "(Undo)" Then
On Error Resume Next
udo_Click
Else
On Error Resume Next
rdo_Click
End If

End Sub

Private Sub Command1_Click()
PopupMenu mnuPopup
End Sub

Private Sub Command2_Click()
Text1.Visible = False
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
colorhtml
Text1.SelStart = 1
Text1.Visible = True
End Sub

Private Sub Command3_Click()
Text1.Visible = False
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
colorhtml
Text1.SelStart = 1
Text1.Visible = True
End Sub

Private Sub copy_Click()
EditCopyProc
End Sub

Private Sub cut_Click()
EditCutProc
End Sub

Private Sub del_Click()
Text1.SelText = ""
End Sub

Private Sub delete_Click()
On Error Resume Next
Kill Dir1.Path & "\" & File1.filename
File1.Refresh
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub exit_Click()
If Text1.Text = "" Then
Unload Form1
End
Else
Dim response
response = MsgBox("This will delete any unsaved code you have. Continue?", vbYesNo + vbQuestion, "Continue?")
If response = vbYes Then
Unload Form1
End
Else
End If
End If
End Sub

Private Sub fbdick_Click()
If fbdick.Checked = True Then
File1.Visible = False
fbdick.Checked = False
Drive1.Visible = False
Dir1.Visible = False
Label2.Visible = False
   Form1.Width = Form1.Width + 20
Else
fbdick.Checked = True
File1.Visible = True
Drive1.Visible = True
Label3.Visible = True
Dir1.Visible = True
   Form1.Width = Form1.Width + 20
End If
End Sub

Private Sub File1_DblClick()
Dim response
response = MsgBox("This will delete any unsaved code you have. Continue?", vbYesNo + vbQuestion, "Continue?")
If response = vbYes Then
On Error Resume Next
NumFile = FileLen(Dir1.Path & "\" & File1.filename)
Open (Dir1.Path & "\" & File1.filename) For Input As #1
MyFile = Input(NumFile, #1)
Text1 = MyFile
Close
Else
End If
End Sub

Private Sub File1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 2 Then
PopupMenu mnu
End If
End Sub

Private Sub Form_Load()
  Set CoolMenuObj = New CoolMenu
   
  Call CoolMenuObj.Install(Me.hwnd, ImageList, True, True)
Dir1.Path = App.Path
Command3.Visible = False
Timer1.Enabled = True

Text3.Visible = False
Text2.Visible = False
  ReDim UndoStack(0)
rdo.Enabled = False
   Text1.Width = Form1.Width - 2650
   
On Error Resume Next
 Dim Åpne3 As String
    Open App.Path & "\option3.txt" For Input As #2
    Input #2, Åpne3
Close
    Text4.Text = Åpne3
    

Text5.Visible = False
Command3_Click
tv1.Nodes.Add.Text = "&nbsp;"
tv1.Nodes.Add.Text = "<!-- *** -->"
tv1.Nodes.Add.Text = "<!DOCTYPE HTML PUBLIC " & Chr(34) & "-//W3C//DTD HTML 4.0//EN" & Chr(34) & " " > ""
tv1.Nodes.Add.Text = "<!DOCTYPE>"
tv1.Nodes.Add.Text = "</TR>"
tv1.Nodes.Add.Text = "<A HREF = " & Chr(34) & " " & Chr(34) & "> </A>"
tv1.Nodes.Add.Text = "<ABBR> </ABBR>"
tv1.Nodes.Add.Text = "<APPLET> </APPLET>"
tv1.Nodes.Add.Text = "<AREA>"
tv1.Nodes.Add.Text = "<B> </B>"
tv1.Nodes.Add.Text = "<BANNER> </BANNER>"
tv1.Nodes.Add.Text = "<BASE>"
tv1.Nodes.Add.Text = "<BASEFONT = " & Chr(34) & " " & Chr(34) & ">"
tv1.Nodes.Add.Text = "<BGCOLOR> </BGCOLOR>"
tv1.Nodes.Add.Text = "<BGSOUND> </BGSOUND>"
tv1.Nodes.Add.Text = "<BIG> </BIG>"
tv1.Nodes.Add.Text = "<BLINK> </BLINK>"
tv1.Nodes.Add.Text = "<BLOCKQUOTE> </BLOCKQUOTE>"
tv1.Nodes.Add.Text = "<BODY> </BODY>"
tv1.Nodes.Add.Text = "<BR>"
tv1.Nodes.Add.Text = "<CAPTION>"
tv1.Nodes.Add.Text = "<CENTER> </CENTER>"
tv1.Nodes.Add.Text = "<CITE> </CITE>"
tv1.Nodes.Add.Text = "<COL>"
tv1.Nodes.Add.Text = "<DEL> </DEL>"
tv1.Nodes.Add.Text = "<DFN> </DFN>"
tv1.Nodes.Add.Text = "<DL> </DL>"
tv1.Nodes.Add.Text = "<FONT FACE = " & Chr(34) & " " & Chr(34) & " SIZE = " & Chr(34) & " " & Chr(34) & " COLOR = " & Chr(34) & " " & Chr(34) & "> </font>"
tv1.Nodes.Add.Text = "<FRAME SCROLLING=YES SRC=" & Chr(34) & "yourpage.htm" & Chr(34) & ">"
tv1.Nodes.Add.Text = "<H?> </H?>"
tv1.Nodes.Add.Text = "<HEAD> </HEAD>"
tv1.Nodes.Add.Text = "<HR>"
tv1.Nodes.Add.Text = "<HTML> </HTML>"
tv1.Nodes.Add.Text = "<I> </I>"
tv1.Nodes.Add.Text = "<INPUT>"
tv1.Nodes.Add.Text = "<INS DATETIME=" & Chr(34) & ":::" & Chr(34) & ">  </INS>"
tv1.Nodes.Add.Text = "<LINK TYPE="" SRC=""></LINK"
tv1.Nodes.Add.Text = "<MARQUEE> </MARQUEE>"
tv1.Nodes.Add.Text = "<META = " & Chr(34) & " " & Chr(34) & ">"
tv1.Nodes.Add.Text = "<NOEMBED> </NOEMBED>"
tv1.Nodes.Add.Text = "<NOSCRIPT> </NOSCRIPT>"
tv1.Nodes.Add.Text = "<P ALIGN  = " & Chr(34) & "CENTER" & Chr(34) & "> </P>"
tv1.Nodes.Add.Text = "<P ALIGN = " & Chr(34) & "LEFT" & Chr(34) & "> </P>"
tv1.Nodes.Add.Text = "<P ALIGN = " & Chr(34) & "RIGHT" & Chr(34) & "> </P>"
tv1.Nodes.Add.Text = "<p>"
tv1.Nodes.Add.Text = "<PRE> </PRE>"
tv1.Nodes.Add.Text = "<Q> </Q>"
tv1.Nodes.Add.Text = "<S> </S>"
tv1.Nodes.Add.Text = "<SAMP> </SAMP>"
tv1.Nodes.Add.Text = "<SCRIPT> </SCRIPTt>"
tv1.Nodes.Add.Text = "<SMALL> </SMALL>"
tv1.Nodes.Add.Text = "<SPACER>"
tv1.Nodes.Add.Text = "<STRONG> </STRONG>"
tv1.Nodes.Add.Text = "<SUB> </SUB>"
tv1.Nodes.Add.Text = "<SUP> </SUP>"
tv1.Nodes.Add.Text = "<TABLE> </TABLE>"
tv1.Nodes.Add.Text = "<TD> </TD>"
tv1.Nodes.Add.Text = "<TI> </TI>"
tv1.Nodes.Add.Text = "<TITLE> </TITLE>"
tv1.Nodes.Add.Text = "<TR>"
tv1.Nodes.Add.Text = "<U> </U>"
DrawLines picLines, Text1
End Sub

Private Sub Form_Resize()
 SSTab1.Height = Form1.Height - StatusBar1.Height - 680
 Text1.Height = Form1.Height - StatusBar1.Height - 680 - Picture3.Height
 File1.Height = Form1.Height - Dir1.Height - Drive1.Height - 1575
 Picture3.Width = Form1.Width
 picLines.Top = Text1.Top + 50
picLines.Height = Text1.Height - 10
 tv1.Height = Form1.Height - StatusBar1.Height - 1800
 Text1.Width = Form1.Width - SSTab1.Width - 200 - picLines.Width
  Picture4.Height = Form1.Height - StatusBar1.Height - 1800 + Label4.Height
  tbToolBar.Width = Form1.Width
End Sub

Private Sub Form_Terminate()
Open "C:\HTMLOG.log" For Append As #1
    Print #1, "***HTM Editor was Exited      "
    Print #1, ""
Close
End Sub

Private Sub HR_Click()
frmHR.Show
End Sub

Private Sub htm_Click()
htm.Checked = True
html.Checked = False
File1.Pattern = "*.htm*"
End Sub

Private Sub html_Click()
htm.Checked = False
html.Checked = True
File1.Pattern = "*.hmtl*"
End Sub

Private Sub import_Click()
On Error Resume Next
Dim html As String
html = InputBox("Enter URL or address", "Enter URL")
Label1.Caption = html
frmBrowser.Show
Unload frmBrowser
Text1.Text = frmBrowser.brwWebBrowser.documentElement.innerHTML
End Sub

Private Sub mnufrme_Click()
frmFrames.Show
End Sub

Private Sub lblfont_Click()

End Sub

Private Sub mnubasicfunctions_Click()
frmBasic.Show
End Sub

Private Sub mnuframse_Click()
Dim response
response = MsgBox("This will delete any unsaved code you have. Continue?", vbYesNo + vbQuestion, "Continue?")
If response = vbYes Then
Unload Me
frmFrames.Show
Else
End If
End Sub

Private Sub mnuimage_Click()
frmImage.Show
End Sub

Private Sub mnuscript_Click()
frmScript.Show
End Sub

Private Sub mnustyle_Click()
frmStyle.Show
End Sub

Private Sub mnutable_Click()
frmTable.Show
End Sub

Private Sub n_Click()
Dim response
response = MsgBox("This will delete any unsaved code you have. Continue?", vbYesNo + vbQuestion, "Continue?")
If response = vbYes Then
Unload Me
frmStartup.Show
Else
End If
End Sub

Private Sub o_Click()
FileOpenProc
sytaxall_Click
End Sub

Private Sub opts_Click()
Form3.Show
End Sub

Private Sub paste_Click()
EditPasteProc
End Sub

Private Sub pcode_Click()
MsgBox "This will print your HTML to your default printer.", vbInformation, "Print File..."
Dim sep, msg
msg = "Your HTML"
sep = "*********************************************************"
sep2 = ""
Printer.Print msg & vbCrLf & _
sep & vbCrLf & sep2 & vbCrLf & Text1 & vbCrLf & _
vbCrLf & vbCrLf & vbCrLf & "DS HTML Editor" & vbCrLf & _
"File Size: " & Len(Text1) & "bytes"
Printer.EndDoc

End Sub

Private Sub rdo_Click()
On Error Resume Next
    UndoStage = UndoStage + 1
    Text1.Text = UndoStack(UndoStage)
    Undoing = False
    
    If sytaxall.Checked = True Then
    Command3_Click
    Else
    End If
End Sub

Private Sub rf_Click()
File1.Refresh
End Sub

Private Sub rname_Click()
frmRename.Show
End Sub

Private Sub s_Click()
On Error GoTo errr:
If Me.Tag = "" Then
With Form1.CommonDialog1
.Filter = "Html pages (*.html)|*.html|All files (*.*)|*.|"
.ShowSave
Form1Text1.SaveFile .filename, 1
Me.Tag = .filename
Form1.Text1.Tag = "no"
End With
Exit Sub
Else
Form1.Text1.SaveFile Me.Tag, 1
Form1.Text1.Tag = "no"
Exit Sub
End If
Exit Sub
errr:
If Err.Number = 32755 Then
Exit Sub
Else
MsgBox Err.Description
End If
End Sub

Private Sub sa_Click()
 Dim strSaveFileName As String
    Dim strDefaultName As String
    If Me.Caption = "HTML Editor- Untitled" Then
        strSaveFileName = GetFileName("Untitled.htm")
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
    Else
        strSaveFileName = GetFileName(strDefaultName)
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
    End If
End Sub

Private Sub selall_Click()
Text1.SelStart = 0
Text1.SelLength = Len(Text1.Text)
End Sub

Private Sub stbar_Click()
If stbar.Checked = True Then
stbar.Checked = False
StatusBar1.Visible = False
   Form1.Width = Form1.Width + 20
Else
stbar.Checked = True
StatusBar1.Visible = True
   Form1.Width = Form1.Width + 20
End If
End Sub

Private Sub sytaxall_Click()
sytaxnone_Click
sytaxnone.Checked = False
sytaxall.Checked = True
Command3_Click
End Sub

Private Sub sytaxnone_Click()
sytaxnone.Checked = True
sytaxall.Checked = False
Text1.TextRTF = Text1.Text
End Sub


Private Sub tbToolBar_ButtonClick(ByVal Button As MSComctlLib.Button)
    On Error Resume Next
    Select Case Button.Key
        Case "New"
            n_Click
        Case "Open"
            FileOpenProc
 If sytaxall.Checked = True Then
 sytaxall_Click
 Else
 End If
        Case "Save"
             Dim strSaveFileName As String
    Dim strDefaultName As String
    If Me.Caption = "HTML Editor- Untitled" Then
        strSaveFileName = GetFileName("Untitled.htm")
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
    Else
        strSaveFileName = GetFileName(strDefaultName)
        If strSaveFileName <> "" Then SaveFileAs (strSaveFileName)
    End If
        Case "Print"
            pcode_Click
        Case "Cut"
            cut_Click
        Case "Copy"
            copy_Click
        Case "Paste"
           paste_Click
        Case "Delete"
            Text1.SelText = ""
        Case "Undo"
            udo_Click
        Case "Redo"
            rdo_Click
        Case "Find"
            yp_Click
        Case "Tab Right"
         PopupMenu mnuPopup
        Case "Help"
        'nothing
    End Select
End Sub

Private Sub Text1_Change()
On Error Resume Next
    ReDim Preserve UndoStack(UBound(UndoStack) + 1)
    UndoStack(UBound(UndoStack)) = Text1.Text
    If Not Undoing Then UndoStage = UndoStage + 1
    
    Label5.Caption = Len(Text1.Text) & " Characters"
    DrawLines picLines, Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If sytaxall.Checked = True Then
On Error Resume Next
If Chr(KeyAscii) = "<" Then
    Text1.SelColor = vbRed
End If
If INtag = True Then
    If Chr(KeyAscii) = " " Then
        If INpropval Then
            Text1.SelColor = vbGreen
        Else
            Text1.SelColor = vbBlue
        End If
    ElseIf Chr(KeyAscii) = """" Then
            Text1.SelColor = &H8000&
    ElseIf Chr(KeyAscii) = ">" Then
            Text1.SelColor = vbRed
    ElseIf Chr(KeyAscii) = "!" Then
            Text1.SelColor = &H8000&
    End If
End If
Else
End If
    DrawLines picLines, Text1
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode & Shift = "1901" Then
    Text1.SelColor = vbBlack
End If
    DrawLines picLines, Text1
End Sub
Private Function INtag() As Boolean
On Error Resume Next
If InStrRev(Text1.Text, "<", Text1.SelStart, vbTextCompare) > InStrRev(Text1.Text, ">", Text1.SelStart, vbTextCompare) Then
INtag = True
End If
End Function
Private Function INpropval() As Boolean
Dim x, y As Long
x = InStrRev(Text1.Text, """", Text1.SelStart, vbTextCompare)
y = InStrRev(Text1.Text, "=", Text1.SelStart, vbTextCompare)
If x > y Then
If InStrRev(Text1.Text, """", x - 1, vbTextCompare) < InStrRev(Text1.Text, "=", x - 1, vbTextCompare) Then INpropval = True
End If
End Function

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    DrawLines picLines, Text1
End Sub

Private Sub time_Click()
Label1.Caption = time
Text1.Text = Text1.Text & " " & Label1.Caption
End Sub

Private Sub Timer2_Timer()
Dim strCopy As String
strCopy = Clipboard.GetText

If strCopy = "" Then
paste.Enabled = False
Else
paste.Enabled = True
End If

If Text1.SelText = "" Then
copy.Enabled = False
cut.Enabled = False
del.Enabled = False

Else
copy.Enabled = True
cut.Enabled = True
del.Enabled = True
End If
End Sub

Private Sub tv1_DblClick()

Text1.SelText = tv1.SelectedItem
End Sub

Private Sub txtfiles_Click()
File1.Pattern = "*.txt*"
End Sub

Private Sub udo_Click()
On Error Resume Next
rdo.Enabled = True
    Undoing = True
    UndoStage = UndoStage - 1
    If UndoStage <= 0 Then UndoStage = 0
    Text1.Text = UndoStack(UndoStage)
    Undoing = False
    
    If sytaxall.Checked = True Then
    Command3_Click
    Else
    End If
End Sub

Private Sub yp_Click()
Dim html As String
html = Text1.Text
Label1.Caption = App.Path & "\temp.html"
Open Label1.Caption For Output As #1
    Print #1, html
    Close #1
frmBrowser.Show
End Sub
Function colorhtml()

    Dim TagregEx, Match, Matches
    Set TagregEx = New RegExp
    TagregEx.Pattern = "<(.)[^> ]*( ){0,1}[^>]*>"
    TagregEx.IgnoreCase = False
    TagregEx.Global = True

    Dim tagPNregEx, Match2, Matches2
    Set tagPNregEx = New RegExp
    tagPNregEx.Pattern = "(\w+ *=) *(\d+|""[^""]+"")"

    tagPNregEx.IgnoreCase = False
    tagPNregEx.Global = True
Dim rtfstart As Long
rtfstart = Text1.SelStart
If Text1.SelLength < 1 Then

Exit Function
End If
    Set Matches = TagregEx.Execute(Text1.SelText)
    For Each Match In Matches
        If Match.Value <> "" Then
            Text1.SelStart = rtfstart + Match.FirstIndex
            Text1.SelLength = Match.Length
            Text1.SelColor = vbRed
            If Match.SubMatches(0) = "!" Then
               Text1.SelColor = &H8000&
               GoTo nextmatch
            ElseIf Match.SubMatches(1) <> " " Then
                GoTo nextmatch
            End If
            Set Matches2 = tagPNregEx.Execute(Match.Value)
            For Each Match2 In Matches2
                If Match2.Value <> "" Then
                    Text1.SelStart = Match.FirstIndex + rtfstart + Match2.FirstIndex
                    Text1.SelLength = Match2.Length
                    Text1.SelColor = &H8000&
                    Text1.SelLength = Len(Match2.SubMatches(0))
                    Text1.SelColor = vbBlue
                End If
            Next
        End If
nextmatch:
    Next
     Label15.Caption = Matches.Count & " Tags"
End Function

Private Sub DrawLines(picTo As PictureBox, RTF As RichTextBox)
Dim iLine As Long, cLine As Long, vLine As Long
iLine = SendMessage(RTF.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
cLine = 1 + RTF.GetLineFromChar(RTF.SelStart)
vLine = SendMessage(RTF.hwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
picTo.Cls

picTo.Font = RTF.Font
picTo.ForeColor = &H8000000C
Dim i As Integer
For i = vLine + 1 To iLine
If i <> cLine Then
picTo.ForeColor = &H8000000C
picTo.FontSize = 10
picTo.Print i
Else
If i = cLine Then
picTo.ForeColor = vbRed
picTo.FontSize = 10
picTo.Print i
End If
End If
Next i
End Sub
