Attribute VB_Name = "Globals"
'@IgnoreModule ConstantNotUsed, ImplicitlyTypedConst, MoveFieldCloserToUsage, EncapsulatePublicField
Option Explicit

' API declarations
#If Mac Then
    Public Declare PtrSafe Function CopyMemory_byVar Lib "libc.dylib" Alias "memmove" (ByRef dest As Any, ByRef src As Any, ByVal size As Long) As LongPtr
#Else
    ' For saving a pointer to the ribbon
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)

    ' For audio recording
    Public Declare PtrSafe Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long
    
    ' VkKeyScan needed to fix the tilde key bug on Macs running Boot Camp
    Public Declare PtrSafe Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
#End If

' UI globals
Public DebateRibbon As IRibbonUI
Public Const USER_FORM_RESIZE_FACTOR As Double = 1.333333

#If Mac Then
    ' Web View is broken on Mac when using the Nav Pane, so default to Draft instead
    Public Const DefaultView As String = "Draft"
#Else
    Public Const DefaultView As String = "Web"
#End If

' Togglebutton state variables
Public AutoOpenFolderToggle As Boolean
Public RecordAudioToggle As Boolean
Public InvisibilityToggle As Boolean
Public UnderlineModeToggle As Boolean
Public ParagraphIntegrityToggle As Boolean
Public UsePilcrowsToggle As Boolean

' Caselist globals
Public Const CASELIST_URL As String = "https://api.opencaselist.com/v1"
Public Const SHARE_URL As String = "https://share.tabroom.com"
Public Const PAPERLESSDEBATE_URL As String = "https://paperlessdebate.com"
Public Const UPDATES_URL As String = "https://update.paperlessdebate.com"
Public Const TABROOM_REGISTER_URL As String = "https://www.tabroom.com/user/login/new_user.mhtml"
Public Const WPM_URL As String = "http://www.readingsoft.com/"

' Paperless globals
Public ActiveSpeechDoc As String

' Audio globals and enums for recording
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1
Public Const SND_FILENAME = &H20000
Public BytesPerSec As Long

' Form UI Constants
Public WHITE As Long
Public BLACK As Long
Public BLUE As Long
Public LIGHT_BLUE As Long
Public GREEN As Long
Public LIGHT_GREEN As Long
Public RED As Long
Public LIGHT_RED As Long
Public ORANGE As Long
Public LIGHT_ORANGE As Long
Public DARK_GRAY As Long

Public Sub InitializeGlobals()
    WHITE = RGB(255, 255, 255) '16777215, &H00FFFFFF&
    BLACK = RGB(0, 0, 0) ' 0, &H00000000&
    BLUE = RGB(64, 92, 121) ' 7953472, &H00795C40&
    LIGHT_BLUE = RGB(114, 142, 171) ' 11243122, &H00AB8E72&
    GREEN = RGB(139, 191, 86) ' 5685131, &H0056BF8B&
    LIGHT_GREEN = RGB(169, 221, 116) ' 7658921, &H0074DDA9&
    RED = RGB(191, 86, 86) ' 5658303, &H005656BF&
    LIGHT_RED = RGB(241, 136, 136) ' 8947953, &H008888F1&
    ORANGE = RGB(191, 139, 86) ' 5671871, &H00568BBF&
    LIGHT_ORANGE = RGB(223, 197, 170) ' 11191775, &H00AAC5DF&
    DARK_GRAY = RGB(169, 169, 169) ' 11119017, &H00A9A9A9&
        
    Globals.ParagraphIntegrityToggle = GetSetting("Verbatim", "Format", "ParagraphIntegrity", True)
    Globals.UsePilcrowsToggle = GetSetting("Verbatim", "Format", "UsePilcrows", True)
End Sub
