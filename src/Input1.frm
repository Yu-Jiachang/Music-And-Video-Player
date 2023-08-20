VERSION 5.00
Begin VB.Form Input1 
   Caption         =   "打开网络文件"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   495
      Left            =   1560
      TabIndex        =   3
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "打开"
      Default         =   -1  'True
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   2400
      Width           =   1455
   End
   Begin VB.TextBox Text 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   4335
   End
   Begin VB.Label Label 
      Caption         =   "请输入要打开的网络文件的地址："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "Input1"
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
