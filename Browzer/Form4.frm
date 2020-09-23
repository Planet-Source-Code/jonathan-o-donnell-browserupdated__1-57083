VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form4 
   Caption         =   "Site Saver"
   ClientHeight    =   3570
   ClientLeft      =   1680
   ClientTop       =   7335
   ClientWidth     =   4290
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form4"
   ScaleHeight     =   3570
   ScaleWidth      =   4290
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   1800
      Top             =   3000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   3120
      TabIndex        =   3
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add Date"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox Text1 
      Height          =   2535
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Site Saver"
      ForeColor       =   &H000000FF&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
On Error Resume Next
Dim FileName As String
CommonDialog1.Filter = "Text Files (*.txt) |*.txt| All Files (*.*) |*.*|"
CommonDialog1.Action = 2
FileName = CommonDialog1.FileName
F = FreeFile
Open FileName For Output As #F '
Print #F, Text1.Text
Close #F
End Sub

Private Sub Command1_Click()
Text1.Text = Text1.Text & Date
End Sub
