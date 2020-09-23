Attribute VB_Name = "Module1"
'Basic function declaration used throughout the program

Option Explicit

'This function gets the current user login from the computer
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

'BitBlt used for drawing
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As _
        Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal _
        hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
Public IsMusicOn As Boolean
Public RetValue As Long

'Used to play wav files
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

'Some constants associates with sndPlaySound
Public Const SND_ASYNC = &H1
Public Const SND_LOOP = &H8
Public Const SND_NODEFAULT = &H2
Public Const SND_SYNC = &H0
Public Const SND_NOSTOP = &H10
Public Const SND_MEMORY = &H4

Public Const max = 10   'max field size
Public curFlame         'next available flame index
Public frameNum         'used with hitAnim and missAnim
'Public frameNum2        'used with nothing
Public frameNum3        'used with radarAnim
Public isDead           'true when ship is dead
Public gameType As Integer  '1-for single and 2-for multi
Public nextShotX As Integer 'computer's next shot X coord
Public nextShotY As Integer 'computer's next shot Y coord

'This structure is used to store each flame info
Public Type flame
    typ As Integer      'was used for type of flame, but later changed to 1 type so now it is used in a diff way (check fireAnim)
    frame As Integer    'current frame the fireAnim is on
    x As Integer        'X coord of flame
    y As Integer        'Y coord of flame
    backImage As PictureBox 'This stores the original field image for x and y
    backImageMask As PictureBox 'This stores the original field image mask for x and y
End Type
Sub main()
frmMain.Show 1
End
End Sub
