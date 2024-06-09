VERSION 5.00
Begin VB.Form Input0 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "打开网络文件"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5655
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   495
      Left            =   2520
      TabIndex        =   3
      Top             =   2400
      Width           =   1500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打开"
      Default         =   -1  'True
      Height          =   495
      Left            =   4080
      TabIndex        =   2
      Top             =   2400
      Width           =   1500
   End
   Begin VB.TextBox Text 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   1560
      Width           =   5415
   End
   Begin VB.Label Label 
      Caption         =   "请输入要打开的网络文件的地址："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2775
   End
End
Attribute VB_Name = "Input0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If Text.Text = "" Then
Command1_Click
Else
Frm.WindowsMediaPlayer.URL = Text.Text
Unload Me
End If
End Sub
