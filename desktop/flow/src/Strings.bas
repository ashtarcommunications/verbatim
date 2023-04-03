Attribute VB_Name = "Strings"
Option Explicit

Public Function OnlySafeChars(ByVal OrigString As String) As String
    Dim lLen As Long
    Dim sAns As String
    Dim lCtr As Long
    Dim sChar As String
    Dim TrimString As String
    
    TrimString = Trim$(OrigString)
    lLen = Len(TrimString)
    For lCtr = 1 To lLen
        sChar = Mid$(TrimString, lCtr, 1)
        If IsSafeChar(Mid$(TrimString, lCtr, 1)) Then
            sAns = sAns & sChar
        End If
    Next
        
    OnlySafeChars = sAns
End Function

Private Function IsSafeChar(ByVal sChr As String) As Boolean
    IsSafeChar = sChr Like "[*0-9A-Za-z -]"
End Function

Public Function OnlyCSV(ByVal OrigString As String) As String
    Dim lLen As Long
    Dim sAns As String
    Dim lCtr As Long
    Dim sChar As String
    Dim TrimString As String
    
    TrimString = Trim$(OrigString)
    lLen = Len(TrimString)
    For lCtr = 1 To lLen
        sChar = Mid$(TrimString, lCtr, 1)
        If IsCSVChar(Mid$(TrimString, lCtr, 1)) Then
            sAns = sAns & sChar
        End If
    Next
        
    OnlyCSV = sAns
End Function

Private Function IsCSVChar(ByVal sChr As String) As Boolean
    IsCSVChar = sChr Like "[*0-9A-Za-z -,\\\/]"
End Function
