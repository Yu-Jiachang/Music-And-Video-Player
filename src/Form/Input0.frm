VERSION 5.00
Begin VB.Form OpenWebFile 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "打开网络文件"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5655
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5655
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Cancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   495
      Left            =   2520
      TabIndex        =   2
      Top             =   2400
      Width           =   1500
   End
   Begin VB.CommandButton Open 
      Caption         =   "打开"
      Default         =   -1  'True
      Height          =   495
      Left            =   4080
      TabIndex        =   3
      Top             =   2400
      Width           =   1500
   End
   Begin VB.TextBox Value 
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   1440
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "请输入要打开的网络文件的地址："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   2775
   End
End
Attribute VB_Name = "OpenWebFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Cancel_Click()
Unload Me
End Sub

Private Sub Open_Click()
If Value.Text = "" Then
Cancel_Click
Else
Frm.MediaPlayer.URL = Value.Text
Cancel_Click
End If
End Sub
