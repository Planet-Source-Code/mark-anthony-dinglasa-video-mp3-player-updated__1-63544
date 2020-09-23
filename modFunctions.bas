Attribute VB_Name = "modFunctions"
'ooooooooooooooooooooooooooooooooooooooooooooooooooooooooo
'xxxxxxxxxxxxxxxxx Tamers Video Player xxxxxxxxxxxxxxxxxxx
'xxxxxxxxxxxxxx Date Created: November 8, 2005  xxxxxxxxxx
'xxxxxxxxxxxxxx Date Finished: November 8, 2005 xxxxxxxxxx
'xxxxxxxxxxxxxx Updated     : December 3,  2005 xxxxxxxxxx
'xxxxxxxxxxxxxx Added       : MP3 Player        xxxxxxxxxx
'ooooooooooooooooooooooooooooooooooooooooooooooooooooooooo

' Note:  This program Uses "Windows Media Player Control",
'        "Windows Common Dialog Control" and "Windows
'        Common Controls 6.0 (SP6).
' Just press [ ctrl + t ] and in the Listbox find all the
' controls that I have mentioned above and check it. Click
' Apply/OK respectively. Done.

'This demo is for beginners Only.

'This is used for Opening and Closing CD/DVD ROMS
Public Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

'This two API's makes up a rounded rectangle control or form
Public Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long

'This is used to make Windows on Top
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'This is used to play sounds like *.wav
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'This is used the Cleaning up the memory used from creating an object
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

'This is used in the replacement of the timer control
'This is more accurate than timer control
Public Declare Function GetTickCount Lib "kernel32" () As Long

'This is only for XP
'This is used to gain XP Themes Look in the Program (Note: app_name.EXE.MANIFEST Required in the same Folder)
Public Declare Function InitCommonControls Lib "Comctl32.dll" () As Long


'Window Position Constants
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOPMOST = -1
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

'Sound Constants
Public Const SND_ASYNC = &H1
Public Const SND_NODEFAULT = &H2

Public VideoDuration As Integer
Public iMute As Integer

'Opens and Close CD/DVD Rom
Public Function CdRom(iOpen As Boolean)
    On Error GoTo Trap
        If iOpen Then
            mciSendString "set cdaudio door open", vbNullString, 0&, 0&
        Else
            mciSendString "set cdaudio door closed", vbNullString, 0&, 0&
        End If
    Exit Function
Trap:
    MsgBox ("There's no CD/DVD Drive Detected !"), vbOKOnly + vbCritical, "CD/DVD Error !"
End Function

'Makes the form always on top of all applications
Public Function TopMost(frm As Form, uSelect As Boolean)
    If uSelect Then
        TopMost = SetWindowPos(frm.hwnd, HWND_TOPMOST, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE)
    Else
        TopMost = SetWindowPos(frm.hwnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_NOSIZE Or SWP_NOMOVE)
    End If
End Function

'Changes the control or a Form to a Rounded shape
Public Function ReShape(uSelect As Boolean, Optional ByVal frm As Form, Optional ByVal Ctl As Control)
    Dim Rgn As Long
        On Error Resume Next
            If uSelect Then
                Rgn = CreateRoundRectRgn(0, 0, frm.Width / Screen.TwipsPerPixelX, frm.Height / Screen.TwipsPerPixelY, 30, 30)
                    SetWindowRgn frm.hwnd, Rgn, True
            Else
                Rgn = CreateRoundRectRgn(0, 0, Ctl.Width / Screen.TwipsPerPixelX, Ctl.Height / Screen.TwipsPerPixelY, 30, 30)
                    SetWindowRgn Ctl.hwnd, Rgn, True
            End If
                        DeleteObject Rgn
End Function

'This function plays a sound (Ex.*.wav)
Public Function MakeAsound(theFilename As String)
    On Error Resume Next
        sndPlaySound theFilename, SND_ASYNC Or SND_NODEFAULT
End Function

'Serves as Timer to update changes when playing Media Player
Public Function SetPosition()
    Dim T1 As Long, T2 As Long
    On Error Resume Next
        T2 = GetTickCount()
            Do
                DoEvents
                    T1 = GetTickCount()
                        If (T1 - T2) >= 1000 Then
                            With frmVideo
                                frmMain.sPosition.Value = CInt(.mPlayer.Controls.currentPosition)
                                frmMain.lblTime.Caption = "Estimate Time: " & .mPlayer.Controls.currentPositionString
                                frmMain.lblDate.Caption = "Video Duration: " & .mPlayer.Controls.currentItem.durationString
                            End With
                                T2 = GetTickCount()
                        End If
            Loop
End Function
