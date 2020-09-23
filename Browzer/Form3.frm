VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form3 
   Caption         =   "IPGetter"
   ClientHeight    =   930
   ClientLeft      =   1440
   ClientTop       =   1575
   ClientWidth     =   5865
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form3"
   ScaleHeight     =   930
   ScaleWidth      =   5865
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3960
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Get IP"
      Height          =   375
      Left            =   4440
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "www."
      Top             =   360
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Site IPGetter"
      ForeColor       =   &H000000FF&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Winsock1.Connect Text1.Text, "80"
DoEvents
End Sub

Private Sub Winsock1_Connect()
Text1.Text = Winsock1.RemoteHostIP
Winsock1.Close
End Sub
