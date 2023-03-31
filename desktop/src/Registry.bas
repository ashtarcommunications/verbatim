Attribute VB_Name = "Registry"
Option Explicit

#If Mac Then
    ' Do Nothing
#Else
Function RegKeyRead(RegKey As String) As String
    Dim WS As Object
    On Error Resume Next
    Set WS = CreateObject("WScript.Shell")
    RegKeyRead = WS.RegRead(RegKey)
    Set WS = Nothing
End Function

Function RegKeyExists(RegKey As String) As Boolean
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

Sub RegKeySave(RegKey As String, Value As String, Optional ValueType As String = "REG_SZ")
    Dim WS As Object
    On Error Resume Next
    Set WS = CreateObject("WScript.Shell")
    WS.RegWrite RegKey, Value, ValueType
    Set WS = Nothing
End Sub
#End If
