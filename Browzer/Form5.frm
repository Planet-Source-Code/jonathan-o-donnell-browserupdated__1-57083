VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Image Viewer"
   ClientHeight    =   5370
   ClientLeft      =   3555
   ClientTop       =   5205
   ClientWidth     =   8295
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form5"
   ScaleHeight     =   5370
   ScaleWidth      =   8295
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   360
      Width           =   2655
   End
   Begin VB.FileListBox File1 
      Height          =   2235
      Left            =   240
      TabIndex        =   2
      Top             =   3000
      Width           =   2655
   End
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.Frame Frame1 
      Caption         =   "Image Viewer"
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
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2895
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   5055
      Left            =   3240
      Stretch         =   -1  'True
      Top             =   240
      Width           =   4935
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
On Error Resume Next
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
On Error Resume Next
    Dir1.Path = Drive1.Drive
End Sub

Private Sub File1_Click()
On Error Resume Next
    Image1.Picture = LoadPicture(Dir1.Path & "\" & File1.FileName)
End Sub
