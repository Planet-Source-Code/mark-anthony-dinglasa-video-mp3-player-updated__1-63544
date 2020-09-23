VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmVideo 
   ClientHeight    =   3660
   ClientLeft      =   3255
   ClientTop       =   1545
   ClientWidth     =   6165
   Icon            =   "frmVideo.frx":0000
   LinkTopic       =   "Form2"
   MouseIcon       =   "frmVideo.frx":0E42
   MousePointer    =   99  'Custom
   ScaleHeight     =   3660
   ScaleWidth      =   6165
   Begin WMPLibCtl.WindowsMediaPlayer mPlayer 
      Height          =   2355
      Left            =   1320
      TabIndex        =   0
      ToolTipText     =   "Double Click to Full Screen/Normal Screen"
      Top             =   600
      Width           =   3480
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
      uiMode          =   "none"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   0   'False
      enableContextMenu=   0   'False
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   6138
      _cy             =   4154
   End
End
Attribute VB_Name = "frmVideo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    VideoDuration = 0 'initialize variable = 0
        Me.Top = Me.Top - 90
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    mPlayer.Controls.stop
    mPlayer.Close
End Sub

Private Sub Form_Resize()
    With mPlayer
        .Move 0, 0, ScaleWidth, ScaleHeight
        .stretchToFit = True
    End With
End Sub

Private Sub mPlayer_Click(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
    Me.ZOrder
End Sub

Private Sub mPlayer_DoubleClick(ByVal nButton As Integer, ByVal nShiftState As Integer, ByVal fX As Long, ByVal fY As Long)
    On Error Resume Next
        If nButton = 1 Then
            mPlayer.fullScreen = True
        End If
End Sub

Private Sub mPlayer_MediaChange(ByVal Item As Object)
    VideoDuration = mPlayer.Controls.currentItem.duration
        If VideoDuration > 1 Then
            frmMain.sPosition.Max = CInt(VideoDuration)
        End If
End Sub
