VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Appearance      =   0  '����
   BackColor       =   &H80000005&
   BorderStyle     =   1  '��u�T�w
   Caption         =   "510208VideoLooker"
   ClientHeight    =   6660
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   9480
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   9480
   StartUpPosition =   3  '�t�ιw�]��
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin WMPLibCtl.WindowsMediaPlayer WMP1 
      Height          =   6675
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9480
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
      stretchToFit    =   -1  'True
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   16722
      _cy             =   11774
   End
   Begin VB.Menu File 
      Caption         =   "�ɮ�(&F)"
      Begin VB.Menu Open 
         Caption         =   "�}��(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu dash 
         Caption         =   "-"
      End
      Begin VB.Menu about 
         Caption         =   "����(&A)"
      End
      Begin VB.Menu exit 
         Caption         =   "���}(&E)"
      End
   End
   Begin VB.Menu playcontrol 
      Caption         =   "����(&P)"
      Begin VB.Menu cmdStop 
         Caption         =   "����(&P)"
      End
      Begin VB.Menu dash2 
         Caption         =   "-"
      End
      Begin VB.Menu full 
         Caption         =   "���ù�(&F)"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdPlay_Click()

End Sub

Private Sub about_Click()
frmAbout.Show
End Sub

Private Sub exit_Click()
t6ext = MsgBox("�T�w�n���}�ܡH", 64 + 1, "���}")
If t6ext = vbOK Then
    End
End If
End Sub

Private Sub Form_Resize()
'WMP1.Width = Form1.Width
'WMP1.Height = Form1.Height
End Sub

Private Sub full_Click()
WMP1.fullScreen = True
End Sub

Private Sub open_Click()
On Error GoTo ErrHandler
'�]�m�L�o���C
CommonDialog1.Filter = "�����ɮ� (*.*)|*.*|MPEG4���Y�榡 (*.mp4)|*.mp4|MPEG�榡���T���Y�� (*.mp3)|*.mp3|WMP�榡�v�T���Y�� (*.wmv)|*.wmv"
'���w�ʬٹL�o���C
CommonDialog1.FilterIndex = 2
'��ܡ����}����ܮءC
CommonDialog1.ShowOpen
'�եΥ��}��󪺹L�{�C
MsgBox "���b�}��" & CommonDialog1.FileName
WMP1.URL = CommonDialog1.FileName
WMP1.play
Exit Sub
ErrHandler:
'�Τ�������������s�C
Exit Sub
End Sub

