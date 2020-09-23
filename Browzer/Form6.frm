VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Explorer"
   ClientHeight    =   4065
   ClientLeft      =   4305
   ClientTop       =   5205
   ClientWidth     =   7080
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form6"
   ScaleHeight     =   4065
   ScaleWidth      =   7080
   Begin VB.FileListBox File1 
      Height          =   3015
      Left            =   240
      TabIndex        =   3
      Top             =   840
      Width           =   3015
   End
   Begin VB.DirListBox Dir1 
      Height          =   3465
      Left            =   3480
      TabIndex        =   2
      Top             =   360
      Width           =   3375
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   3015
   End
   Begin VB.Frame Frame1 
      Caption         =   "Explorer"
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
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6855
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

 
