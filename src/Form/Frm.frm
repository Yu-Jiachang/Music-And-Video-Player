VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Frm 
   Caption         =   "视频播放器"
   ClientHeight    =   5160
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13470
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   13470
   StartUpPosition =   3  '窗口缺省
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   12960
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "打开本地文件"
      InitDir         =   "C:\Users\Yu\Videos"
   End
   Begin WMPLibCtl.WindowsMediaPlayer MediaPlayer 
      Height          =   5055
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   1
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   23733
      _cy             =   8916
   End
   Begin VB.Menu File 
      Caption         =   "文件(&F)"
      Begin VB.Menu OpenFile 
         Caption         =   "打开本地文件"
         Shortcut        =   ^O
      End
      Begin VB.Menu Open_WebFile 
         Caption         =   "打开网络文件"
         Shortcut        =   ^W
      End
      Begin VB.Menu Stop 
         Caption         =   "停止"
         Shortcut        =   ^S
      End
      Begin VB.Menu Exit 
         Caption         =   "退出"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu About 
         Caption         =   "关于"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Frm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()
On Error Resume Next
MediaPlayer.Height = Me.ScaleHeight
MediaPlayer.Width = Me.ScaleWidth
End Sub

Private Sub OpenFile_Click()
On Error GoTo Cancel
CommonDialog.ShowOpen
MediaPlayer.URL = CommonDialog.FileName
Cancel:
End Sub

Private Sub Open_WebFile_Click()
OpenWebFile.Show vbModal
End Sub

Private Sub Stop_Click()
MediaPlayer.Close
End Sub

Private Sub Exit_Click()
End
End Sub
