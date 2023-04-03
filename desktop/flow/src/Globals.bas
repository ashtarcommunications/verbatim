Attribute VB_Name = "Globals"
'@IgnoreModule VariableNotUsed, ConstantNotUsed, ImplicitlyTypedConst, MoveFieldCloserToUsage, EncapsulatePublicField
Option Explicit

Public Const VERSION As String = "1.0.0"

' API declarations
#If Mac Then
    Public Declare PtrSafe Function CopyMemory_byVar Lib "libc.dylib" Alias "memmove" (ByRef dest As Any, ByRef src As Any, ByVal size As Long) As LongPtr
#Else
    ' For saving a pointer to the ribbon
    Public Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)
#End If

' UI globals
Public DebateRibbon As IRibbonUI
Public Const USER_FORM_RESIZE_FACTOR As Double = 1.333333

' Togglebutton state variables
Public InsertModeToggle As Boolean

' Form UI Constants
Public WHITE As Long
Public YELLOW As Long
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
    YELLOW = RGB(255, 255, 0) '65535, &H00FFFF00&
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
End Sub
