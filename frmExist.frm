VERSION 5.00
Begin VB.Form frmExist 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "File Exists!"
   ClientHeight    =   1260
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
   ForeColor       =   &H8000000F&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1260
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   120
      Picture         =   "frmExist.frx":0000
      ScaleHeight     =   495
      ScaleWidth      =   495
      TabIndex        =   4
      Top             =   360
      Width           =   495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Resume"
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Overwrite"
      Height          =   375
      Left            =   1200
      TabIndex        =   1
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "File Already Exists What Would You Like To Do?"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   840
      TabIndex        =   0
      Top             =   240
      Width           =   3945
   End
End
Attribute VB_Name = "frmExist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Kill FilePathName
RESUMEFILE = False
Main.CloseSocket
Main.Winsock.Connect strSvrURL, 80
Unload Me
Unload frmSave
End Sub

Private Sub Command2_Click()
RESUMEFILE = True
Main.CloseSocket
FileLength = FileLen(FilePathName)
Main.Winsock.Connect strSvrURL, 80
Unload Me
Unload frmSave
End Sub

Private Sub Command3_Click()
Unload Me
Unload frmSave
End Sub

