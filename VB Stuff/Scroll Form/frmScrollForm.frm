VERSION 5.00
Begin VB.Form frmScrollForm 
   Caption         =   "Form1"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4245
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   4245
   StartUpPosition =   3  'Windows Default
   Begin VB.FileListBox File1 
      Height          =   2040
      Left            =   3600
      TabIndex        =   6
      Top             =   4200
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   2115
      Left            =   2160
      TabIndex        =   5
      Top             =   4200
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   720
      TabIndex        =   4
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   855
      Left            =   4560
      TabIndex        =   3
      Text            =   "Text Box"
      Top             =   240
      Width           =   3495
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   6240
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   495
      Left            =   6960
      TabIndex        =   1
      Top             =   6240
      Width           =   1215
   End
   Begin VB.PictureBox picForm 
      AutoRedraw      =   -1  'True
      Height          =   3735
      Left            =   0
      Picture         =   "frmScrollForm.frx":0000
      ScaleHeight     =   3675
      ScaleWidth      =   4155
      TabIndex        =   0
      Top             =   0
      Width           =   4215
      Begin VB.Image Image1 
         Height          =   480
         Index           =   0
         Left            =   840
         Picture         =   "frmScrollForm.frx":36D2
         Top             =   2520
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmScrollForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command3_Click()

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Dir1_Change()
    File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

' Just some notes from me. To add non hwnd controls
' like shapes they must be be in a container like a
' picturebox with Autoredraw = True and ScaleMode =
' Pixel until the AddScrollBar routine has completed.
' This is to allow correct determination of where the
' left and bottom edges of all controls are so that
' when scrolling you do not move the controls completely
' off the viewport.  jgd 01/19/2002
Private Sub Form_Load()
    picForm.ScaleMode = vbPixels
    picForm.BorderStyle = 0
    picForm.AutoRedraw = True
    AddScrollBar Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DestroyScrollBar
End Sub
