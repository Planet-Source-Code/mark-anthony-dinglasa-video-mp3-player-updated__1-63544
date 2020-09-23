VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tamers Video Player !"
   ClientHeight    =   2115
   ClientLeft      =   3255
   ClientTop       =   5625
   ClientWidth     =   6165
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMain.frx":1CFA
   MousePointer    =   99  'Custom
   ScaleHeight     =   2115
   ScaleWidth      =   6165
   Begin MSComDlg.CommonDialog cBrowse 
      Left            =   1200
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdStop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   495
      Left            =   4930
      MouseIcon       =   "frmMain.frx":2004
      MousePointer    =   99  'Custom
      TabIndex        =   5
      ToolTipText     =   "Click to stop playing !"
      Top             =   1060
      Width           =   1095
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   495
      Left            =   3730
      MouseIcon       =   "frmMain.frx":230E
      MousePointer    =   99  'Custom
      TabIndex        =   4
      ToolTipText     =   "Click to Pause and Resume playing !"
      Top             =   1060
      Width           =   1095
   End
   Begin VB.CommandButton cmdPlay 
      Caption         =   "Play"
      Enabled         =   0   'False
      Height          =   495
      Left            =   2530
      MouseIcon       =   "frmMain.frx":2618
      MousePointer    =   99  'Custom
      TabIndex        =   3
      ToolTipText     =   "Click to Play the File you Selected !"
      Top             =   1060
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   495
      Left            =   1330
      MouseIcon       =   "frmMain.frx":2922
      MousePointer    =   99  'Custom
      TabIndex        =   2
      ToolTipText     =   "Click to Open a File to Play !"
      Top             =   1060
      Width           =   1095
   End
   Begin VB.CommandButton cmdCdRom 
      Caption         =   "CD Open"
      Height          =   495
      Left            =   130
      MouseIcon       =   "frmMain.frx":2C2C
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Click to Open and Close CD-ROM !"
      Top             =   1060
      Width           =   1095
   End
   Begin MSComctlLib.Slider sVolume 
      Height          =   255
      Left            =   4440
      TabIndex        =   0
      ToolTipText     =   "Change Volume Here !"
      Top             =   1800
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      MousePointer    =   99
      MouseIcon       =   "frmMain.frx":2F36
      LargeChange     =   30
      Min             =   4
      Max             =   200
      SelStart        =   5
      TickStyle       =   3
      Value           =   5
   End
   Begin MSComctlLib.Slider sPosition 
      Height          =   375
      Left            =   0
      TabIndex        =   6
      TabStop         =   0   'False
      ToolTipText     =   "Click here to move frames Backward/Forward !"
      Top             =   520
      Width           =   6180
      _ExtentX        =   10901
      _ExtentY        =   661
      _Version        =   393216
      MousePointer    =   99
      MouseIcon       =   "frmMain.frx":3250
      LargeChange     =   100
      SmallChange     =   10
      Max             =   1
      TickStyle       =   3
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   120
      MouseIcon       =   "frmMain.frx":356A
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":3874
      Stretch         =   -1  'True
      ToolTipText     =   "Click to Open MP3 Window"
      Top             =   1800
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   255
      Left            =   120
      Picture         =   "frmMain.frx":46B6
      Stretch         =   -1  'True
      Top             =   150
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   3960
      MouseIcon       =   "frmMain.frx":63B0
      MousePointer    =   99  'Custom
      Picture         =   "frmMain.frx":66BA
      Stretch         =   -1  'True
      ToolTipText     =   "Click to Mute / Normal Sounds !"
      Top             =   1800
      Width           =   495
   End
   Begin VB.Label chkTop 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Click To Top !"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   4460
      MouseIcon       =   "frmMain.frx":83B4
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   150
      Width           =   1470
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   3960
      Picture         =   "frmMain.frx":86BE
      Stretch         =   -1  'True
      Top             =   150
      Width           =   495
   End
   Begin VB.Label lblDate 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   600
      TabIndex        =   8
      Top             =   150
      Width           =   510
   End
   Begin VB.Label lblTime 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   240
      Left            =   600
      TabIndex        =   7
      Top             =   1800
      Width           =   540
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   4
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   -2760
      Top             =   -135
      Width           =   9135
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   4
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   -120
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
      Width           =   6735
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
    
Private Sub chkTop_Click()
    MakeAsound App.Path & "\reload1_44.wav"
    If chkTop.Caption = "Click To Top !" Then
        TopMost Me, True
        TopMost frmVideo, True
        TopMost frmMP3, True
            chkTop.ForeColor = vbYellow
            chkTop.Caption = "Top Enabled !"
                lblDate.ForeColor = vbYellow
                lblTime.ForeColor = vbYellow
                chkTop.ForeColor = vbYellow
    Else
        TopMost Me, False
        TopMost frmVideo, False
        TopMost frmMP3, False
            chkTop.ForeColor = &HFFC0C0
            chkTop.Caption = "Click To Top !"
                lblDate.ForeColor = &HFFC0C0
                lblTime.ForeColor = &HFFC0C0
                chkTop.ForeColor = &HFFC0C0
    End If
End Sub

Private Sub chkTop_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    With chkTop
        .Move .Left - 200
    End With
End Sub

Private Sub chkTop_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    With chkTop
        .Move 4440
    End With
End Sub

Private Sub cmdCdRom_Click()
    MakeAsound App.Path & "\reload1_44.wav"
        If cmdCdRom.Caption = "CD Open" Then
            CdRom True
            cmdCdRom.Caption = "CD Close"
        Else
            CdRom False
            cmdCdRom.Caption = "CD Open"
        End If
End Sub

Private Sub cmdOpen_Click()
    MakeAsound App.Path & "\reload1_44.wav"
        With cBrowse
            .Filter = "DAT Files(*.dat)|*.dat|AVI Files(*.avi)|*.avi|WMV File(*.wmv)|*.wmv|MPEG Files(*.mpg)|*.mpg|All Files (*.*)|*.*"
            .FilterIndex = 0
            .FileName = ""
            .ShowOpen
        End With
                If Not cBrowse.FileName = "" Then
                    Call cmdPlay_Click
                End If
End Sub

Private Sub cmdPause_Click()
Dim sFilter As String
    If cmdPause.Caption = "Pause" Then
        With frmVideo
            .Hide
            .mPlayer.Controls.pause
        End With
            cmdPause.Caption = "Resume"
    Else
        sFilter = Right$(cBrowse.FileName, 3)
            Select Case LCase(sFilter)
                Case "mp3", "wav", "mid", "idi", ""
                    frmVideo.Hide
                Case Else
                    frmVideo.Show
            End Select
                    frmVideo.mPlayer.Controls.play
                    cmdPause.Caption = "Pause"
    End If
                        MakeAsound App.Path & "\reload1_44.wav"
End Sub

Private Sub cmdPlay_Click()
    Dim datFilter As String
        On Error Resume Next
            datFilter = Right$(cBrowse.FileName, 3)
                Select Case LCase(datFilter)
                    Case "mp3", "wav", "mid", "idi"
                        frmVideo.Hide
                    Case Else
                        frmVideo.Show
                    End Select
                            PlayVideo
                                SetPosition
                                    MakeAsound App.Path & "\reload1_44.wav"
End Sub

Private Sub cmdStop_Click()
    With frmVideo
        .mPlayer.Controls.stop
    End With
        frmVideo.Hide
            cmdPause.Enabled = False
            cmdStop.Enabled = False
            cmdPlay.Enabled = False
                MakeAsound App.Path & "\reload1_44.wav"
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    If App.PrevInstance = True Then End
        lblDate.Caption = Format(Date, "Long Date")
        lblTime.Caption = Format(Time, "Long Time")
            sVolume.Value = GetSetting(App.Title, "Settings", "Volume", 50)
                With frmVideo.mPlayer.settings
                    .volume = sVolume.Value
                    .balance = True
                    .enableErrorDialogs = True
                End With
                        iMute = 1
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
        Unload frmMP3
        Unload frmVideo
            SaveSetting App.Title, "Settings", "Volume", sVolume.Value
                End
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        If Not Me.Height = 2640 Or Not Me.Width = 6285 Then
            Me.Height = 2640
            Me.Width = 6285
        End If
End Sub

Private Sub Image2_Click()
    MakeAsound App.Path & "\reload1_44.wav"
        With frmVideo
            If iMute = 1 Then
                .mPlayer.settings.mute = True
                    Image2.Picture = LoadPicture(App.Path & "\Mute.ico")
                        iMute = 0
            Else
                .mPlayer.settings.mute = False
                    Image2.Picture = LoadPicture(App.Path & "\Sound.ico")
                        iMute = 1
            End If
        End With
End Sub

Private Sub Image4_Click()
    If frmMP3.Visible = False Then
        frmMP3.Show
            MakeAsound App.Path & "\reload1_44.wav"
    End If
End Sub

Private Sub sPosition_Change()
     Dim i As Integer
        If frmMP3.Visible = True Then
            If sPosition.Value >= sPosition.Max - 1 Then
                If frmMP3.Playlist.ListCount <> 0 Then
                    On Error Resume Next
                        Randomize
                            i = CInt(Rnd * frmMP3.Playlist.ListCount)
                                frmMP3.Playlist.Text = frmMP3.Playlist.List(i)
                                    With frmVideo
                                        .mPlayer.Controls.stop
                                        .mPlayer.URL = frmMP3.dirSelect.Path & "\" & frmMP3.Playlist.Text & ".mp3"
                                        .mPlayer.Controls.play
                                            frmMP3.lblCurrent.Caption = "Playing - " & UCase(frmMP3.Playlist.Text) & ".MP3"
                                                cmdStop.Enabled = True
                                                cmdPause.Enabled = True
                                    End With
                                                    SetPosition
                End If
            End If
        End If
End Sub

Private Sub sPosition_Click()
    With frmVideo.mPlayer.Controls
        .currentPosition = sPosition.Value
    End With
End Sub

Private Sub sVolume_Change()
    With frmVideo.mPlayer
        .settings.volume = sVolume.Value
    End With
End Sub

Private Function PlayVideo() As Long
    With frmVideo
        .Caption = "Tamers Video Player: " & cBrowse.FileName
            .mPlayer.Controls.stop
            .mPlayer.URL = cBrowse.FileName
            .mPlayer.Controls.play
    End With
                cmdPause.Enabled = True
                cmdStop.Enabled = True
                cmdPlay.Enabled = True
End Function

