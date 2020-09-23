VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "msinet.ocx"
Begin VB.Form Form2 
   Caption         =   "HTML Getter"
   ClientHeight    =   6000
   ClientLeft      =   8175
   ClientTop       =   4575
   ClientWidth     =   5415
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   6000
   ScaleWidth      =   5415
   Begin VB.TextBox txtSrc 
      Height          =   4815
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   960
      Width           =   4935
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4560
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton cmdGet 
      Caption         =   "Get Source"
      Height          =   375
      Left            =   3960
      TabIndex        =   2
      Top             =   360
      Width           =   1215
   End
   Begin VB.TextBox txtHst 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Text            =   "http://"
      Top             =   360
      Width           =   3615
   End
   Begin VB.Frame Frame1 
      Caption         =   "Get Site Source"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdGet_Click()
On Error Resume Next
txtSrc.Text = Inet1.OpenURL(txtHst.Text)
End Sub





