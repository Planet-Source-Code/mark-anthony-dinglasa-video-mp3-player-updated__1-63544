VERSION 5.00
Begin VB.Form frmMP3 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   Caption         =   "Tamers MP3:"
   ClientHeight    =   3705
   ClientLeft      =   3255
   ClientTop       =   1545
   ClientWidth     =   6165
   Icon            =   "frmMP3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmMP3.frx":08CA
   MousePointer    =   99  'Custom
   ScaleHeight     =   3705
   ScaleWidth      =   6165
   Begin VB.Frame fSelect 
      BackColor       =   &H00400000&
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   520
      MouseIcon       =   "frmMP3.frx":0BD4
      MousePointer    =   99  'Custom
      TabIndex        =   2
      Top             =   540
      Visible         =   0   'False
      Width           =   5055
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel"
         Height          =   375
         Left            =   3000
         MouseIcon       =   "frmMP3.frx":0EDE
         MousePointer    =   99  'Custom
         TabIndex        =   9
         Top             =   2220
         Width           =   1455
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "OK"
         Height          =   375
         Left            =   600
         MouseIcon       =   "frmMP3.frx":11E8
         MousePointer    =   99  'Custom
         TabIndex        =   5
         Top             =   2220
         Width           =   1455
      End
      Begin VB.DirListBox dirSelect 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   1665
         Left            =   120
         MouseIcon       =   "frmMP3.frx":14F2
         MousePointer    =   99  'Custom
         TabIndex        =   4
         Top             =   480
         Width           =   4815
      End
      Begin VB.DriveListBox driveSelect 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   120
         MouseIcon       =   "frmMP3.frx":17FC
         TabIndex        =   3
         Top             =   120
         Width           =   4815
      End
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   365
      Left            =   120
      TabIndex        =   13
      Text            =   "Search Songs Here !"
      ToolTipText     =   "Type here to search a song !"
      Top             =   1000
      Width           =   5895
   End
   Begin VB.FileListBox FileListBox1 
      Height          =   285
      Left            =   5040
      Pattern         =   "*.mp3"
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Playlist"
      Height          =   375
      Left            =   1560
      MouseIcon       =   "frmMP3.frx":1B06
      MousePointer    =   99  'Custom
      TabIndex        =   8
      ToolTipText     =   "Click to clear playlist !"
      Top             =   600
      Width           =   1335
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "Select Folder"
      Height          =   375
      Left            =   129
      MouseIcon       =   "frmMP3.frx":1E10
      MousePointer    =   99  'Custom
      TabIndex        =   1
      ToolTipText     =   "Click to select source mp3 files !"
      Top             =   600
      Width           =   1455
   End
   Begin VB.ListBox Playlist 
      Appearance      =   0  'Flat
      BackColor       =   &H00400000&
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
      Height          =   1710
      Left            =   129
      MouseIcon       =   "frmMP3.frx":211A
      MousePointer    =   99  'Custom
      TabIndex        =   0
      ToolTipText     =   "Select a song to play !"
      Top             =   1400
      Width           =   5895
   End
   Begin VB.Label lblSource 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Playing - "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   240
      Left            =   3120
      TabIndex        =   12
      Top             =   650
      Width           =   990
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   3960
      Picture         =   "frmMP3.frx":2424
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   240
   End
   Begin VB.Label lblcount 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Files"
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
      Left            =   4320
      TabIndex        =   10
      Top             =   3360
      Width           =   525
   End
   Begin VB.Label lblFiles 
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
      TabIndex        =   7
      Top             =   3360
      Width           =   510
   End
   Begin VB.Image Image1 
      Height          =   240
      Left            =   240
      Picture         =   "frmMP3.frx":3266
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   240
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   240
      Picture         =   "frmMP3.frx":40A8
      Stretch         =   -1  'True
      Top             =   120
      Width           =   240
   End
   Begin VB.Label lblCurrent 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Playing - "
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
      TabIndex        =   6
      Top             =   120
      Width           =   990
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   4
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   -120
      Top             =   3240
      Width           =   9135
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FF8080&
      BorderWidth     =   4
      FillColor       =   &H00400000&
      FillStyle       =   0  'Solid
      Height          =   1095
      Left            =   -120
      Top             =   -600
      Width           =   9135
   End
End
Attribute VB_Name = "frmMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
    fSelect.Visible = False
End Sub

Private Sub cmdClear_Click()
    Playlist.Clear
End Sub

Private Sub cmdOK_Click()
    Dim sFilter As String, iList As Integer
        fSelect.Visible = False
            If FileListBox1.ListCount <> 0 Then
                Playlist.Clear
                    For iList = 0 To FileListBox1.ListCount - 1
                        DoEvents
                            sFilter = Left$(FileListBox1.List(iList), Len(FileListBox1.List(iList)) - 4)
                                Playlist.AddItem UCase(sFilter)
                    Next
            End If
                                    lblFiles.Caption = Playlist.ListCount - 1 & " files in the Playlist"
                                    lblSource.Caption = dirSelect.Path
End Sub

Private Sub cmdSelect_Click()
    fSelect.Visible = True
    ReShape False, , fSelect
End Sub

Private Sub dirSelect_Change()
    FileListBox1.Path = dirSelect.Path
End Sub

Private Sub dirSelect_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        dirSelect.Path = dirSelect.List(dirSelect.ListIndex)
    End If
End Sub

Private Sub driveSelect_Change()
    On Error Resume Next
        dirSelect.Path = driveSelect.Drive
End Sub

Private Sub Form_Initialize()
    InitCommonControls
End Sub

Private Sub Form_Load()
    On Error Resume Next
        driveSelect.Drive = GetSetting(App.Title, "MP3", "pDrive", App.Path)
        dirSelect.Path = GetSetting(App.Title, "MP3", "pListBox", App.Path)
            Call cmdOK_Click
                Me.Top = Me.Top - 100
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    With frmVideo
        .mPlayer.Controls.stop
        .mPlayer.Close
    End With
            SaveSetting App.Title, "MP3", "pDrive", driveSelect.Drive
            SaveSetting App.Title, "MP3", "pListBox", dirSelect.Path
End Sub

Private Sub Form_Resize()
    On Error Resume Next
        If Not Me.Width = 6285 Or Not Me.Height = 4215 Then
            Me.Width = 6285
            Me.Height = 4215
        End If
End Sub

Private Sub Playlist_Click()
    If Playlist.ListCount <> 0 Then
        Me.Caption = "Tamers MP3: " & Playlist.Text
            lblcount.Caption = Playlist.ListIndex & " of " & Playlist.ListCount - 1 & " Files"
    End If
End Sub

Private Sub Playlist_DblClick()
    If Playlist.ListCount <> 0 Then
        On Error Resume Next
            With frmVideo
                .mPlayer.Controls.stop
                .mPlayer.URL = dirSelect.Path & "\" & Playlist.Text & ".mp3"
                .mPlayer.Controls.play
                    lblCurrent.Caption = "Playing - " & Playlist.Text & ".MP3"
                        frmMain.cmdStop.Enabled = True
                        frmMain.cmdPause.Enabled = True
                            SetPosition
            End With
    End If
End Sub

Private Sub Playlist_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Playlist_DblClick
    End If
End Sub

Private Sub txtSearch_Change()
    Dim i As Integer, sWord As String, sMatch As String
        sWord = txtSearch.Text
            For i = 0 To Playlist.ListCount - 1
                DoEvents
                    sMatch = Left$(Playlist.List(i), Len(txtSearch.Text))
                        If txtSearch.Text = "" Then
                            With Playlist
                                .Text = .List(0)
                            End With
                                Exit For
                        ElseIf sMatch = UCase(txtSearch.Text) Then
                            With Playlist
                                .Text = .List(i)
                            End With
                                Exit For
                        End If
            Next
End Sub

Private Sub txtSearch_GotFocus()
    With txtSearch
        .SelStart = 0
        .SelLength = Len(.Text)
        .SetFocus
    End With
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call txtSearch_Change
        Call Playlist_DblClick
    End If
End Sub

Private Sub txtSearch_LostFocus()
    txtSearch.Text = "Search Songs Here !"
End Sub
