VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form Form1 
   Caption         =   "JizZy'S~Browser"
   ClientHeight    =   9360
   ClientLeft      =   1560
   ClientTop       =   1455
   ClientWidth     =   12120
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9360
   ScaleWidth      =   12120
   Begin VB.CommandButton Command13 
      Caption         =   "Explorer"
      Height          =   375
      Left            =   10560
      TabIndex        =   18
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command12 
      Caption         =   "View Image"
      Height          =   375
      Left            =   8520
      TabIndex        =   17
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Site Saver"
      Height          =   375
      Left            =   6360
      TabIndex        =   16
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Yahoo"
      Height          =   375
      Left            =   4320
      TabIndex        =   15
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command9 
      Caption         =   " Get  Site IP"
      Height          =   375
      Left            =   2280
      TabIndex        =   14
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton Command8 
      Caption         =   "SrcGetter"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Frame Frame4 
      Height          =   735
      Left            =   120
      TabIndex        =   12
      Top             =   8520
      Width           =   11895
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Exit"
      Height          =   375
      Left            =   10440
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   10
      Text            =   "http://"
      Top             =   1080
      Width           =   11655
   End
   Begin VB.Frame Frame3 
      Caption         =   "Site URL here"
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   11895
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Msn Search"
      Height          =   375
      Left            =   8640
      TabIndex        =   8
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Home"
      Height          =   375
      Left            =   6960
      TabIndex        =   7
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next"
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Back"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   240
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser wb 
      Height          =   6615
      Left            =   240
      TabIndex        =   2
      Top             =   1800
      Width           =   11655
      ExtentX         =   20558
      ExtentY         =   11668
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Frame Frame2 
      Caption         =   "Display"
      ForeColor       =   &H000000FF&
      Height          =   6975
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   11895
   End
   Begin VB.Frame Frame1 
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next ' if theres an error ignore it '
URL = "" & Combo1.Text
wb.Navigate URL   'When command1 is clicked it navigates to the site typed in combo1.text '
End Sub

Private Sub Command10_Click()
On Error Resume Next
URL = "http://www.yahoo.com"
wb.Navigate URL  'When Command10 is clicked it navigates to the yahoo homepage '
End Sub

Private Sub Command11_Click()
Form4.Show 'When Command11 is clicked it will show form4 like a link so to speak '
End Sub

Private Sub Command12_Click()
Form5.Show
End Sub

Private Sub Command13_Click()
Form6.Show
End Sub

Private Sub Command2_Click()
On Error Resume Next
wb.GoBack  'Tells the wb (webrowser) to go back
End Sub

Private Sub Command3_Click()
On Error Resume Next
wb.GoForward  'Tells the wb (webrowser) to go Forward
End Sub

Private Sub Command4_Click()
On Error Resume Next
wb.Refresh   'Tells the wb (webrowser) to Refresh
End Sub

Private Sub Command5_Click()
On Error Resume Next
wb.GoHome  'Tells the wb (webrowser) to go home
End Sub

Private Sub Command6_Click()
On Error Resume Next
wb.GoSearch   'Tells the wb (webrowser) to go to the msn serch page
End Sub

Private Sub Command7_Click()
On Error Resume Next
MsgBox "Thanks For Usin JizZy's Browser", vbOKOnly, "Exit Browser" '
End  ' Ends Program '
End Sub

Private Sub Command8_Click()
Form2.Show
End Sub

Private Sub Command9_Click()
Form3.Show
End Sub
