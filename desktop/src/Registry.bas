Attribute VB_Name = "Registry"
'@IgnoreModule ProcedureNotUsed
Option Explicit

#If Mac Then
    ' Do Nothing
#Else
Public Function RegKeyRead(ByRef RegKey As String) As String
    Dim WS As Object
    On Error Resume Next
    Set WS = CreateObject("WScript.Shell")
    RegKeyRead = WS.RegRead(RegKey)
    Set WS = Nothing
    On Error GoTo 0
End Function

Public Function RegKeyExists(ByRef RegKey As String) As Boolean
    Dim WS As Object
    On Error GoTo Handler
    Set WS = CreateObject("WScript.Shell")
    WS.RegRead RegKey
    RegKeyExists = True
    Set WS = Nothing
    Exit Function
  
Handler:
    ' If key isn't found, it will throw an error
    RegKeyExists = False
    Set WS = Nothing
End Function

Public Sub RegKeySave(ByRef RegKey As String, ByRef Value As String, Optional ByRef ValueType As String = "REG_SZ")
    Dim WS As Object
    On Error Resume Next
    Set WS = CreateObject("WScript.Shell")
    WS.RegWrite RegKey, Value, ValueType
    Set WS = Nothing
    On Error GoTo 0
End Sub
#End If
