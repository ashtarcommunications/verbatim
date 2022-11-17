Attribute VB_Name = "Globals"
' Windows API declarations
#If Win64 Then
    ' For saving a pointer to the ribbon
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)

    ' For audio recording
    Public Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    
    ' VkKeyScan needed to fix the tilde key bug on Macs running Boot Camp
    Public Declare PtrSafe Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
    
    'ShellExecute needed to launch installer package after updates
    Public Declare PtrSafe Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
#End If

Public IsMac As Boolean

' UI globals
Public DebateRibbon As IRibbonUI
Public Const USER_FORM_RESIZE_FACTOR As Double = 1.333333

#If Mac Then
    ' Web View is broken on Mac when using the Nav Pane, so default to Draft instead
    Public Const DefaultView As String = "Draft"
#Else
    Public Const DefaultView As String = "Web"
#End If

'Togglebutton state variables
Public AutoOpenFolderToggle As Boolean
Public AutoCoauthoringToggle As Boolean
Public RecordAudioToggle As Boolean
Public InvisibilityToggle As Boolean
Public UnderlineModeToggle As Boolean

' Caselist globals
Public Const CASELIST_URL As String = "https://api.opencaselist.com/v1"
Public Const SHARE_URL As String = "https://share.tabroom.com"
Public Const MOCK_ROUNDS As String = "https://run.mocky.io/v3/be382c53-e49c-4de4-99b6-44ba5e6a3e7c"
 
' Paperless globals
Public ActiveSpeechDoc As String

' Audio globals and enums for recording
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000

Public Enum BitsPerSec
    Bits16 = 16
    Bits8 = 8
End Enum

Public Enum SamplesPerSec
    Samples8000 = 8000
    Samples11025 = 11025
    Samples12000 = 12000
    Samples16000 = 16000
    Samples22050 = 22050
    Samples24000 = 24000
    Samples32000 = 32000
    Samples44100 = 44100
    Samples48000 = 48000
End Enum

Public Enum Channels
    Mono = 1
    Stereo = 2
End Enum

' Form UI Constants
Public WHITE As Long
Public BLACK As Long
Public BLUE As Long
Public LIGHT_BLUE As Long
Public GREEN As Long
Public LIGHT_GREEN As Long
Public RED As Long
Public LIGHT_RED As Long
Public DARK_GRAY As Long

Sub InitializeGlobals()
    WHITE = RGB(255, 255, 255) '16777215, &H00FFFFFF&
    BLACK = RGB(0, 0, 0) ' 0, &H00000000&
    BLUE = RGB(64, 92, 121) ' 7953472, &H00795C40&
    LIGHT_BLUE = RGB(114, 142, 171) ' 11243122, &H00AB8E72&
    GREEN = RGB(139, 191, 86) ' 5685131, &H0056BF8B&
    LIGHT_GREEN = RGB(169, 221, 116) ' 7658921, &H0074DDA9&
    RED = RGB(191, 86, 86) ' 5658303, &H005656BF&
    LIGHT_RED = RGB(241, 136, 136) ' 8947953, &H008888F1&
    DARK_GRAY = RGB(169, 169, 169) ' 11119017, &H00A9A9A9&
    
    
End Sub


