Attribute VB_Name = "Strings"
Option Explicit

Public Function OnlyAlphaNumericChars(ByVal OrigString As String) As String
    Dim lLen As Long
    Dim sAns As String
    Dim lCtr As Long
    Dim sChar As String
    
    OrigString = Trim(OrigString)
    lLen = Len(OrigString)
    For lCtr = 1 To lLen
        sChar = Mid(OrigString, lCtr, 1)
        If IsAlphaNumeric(Mid(OrigString, lCtr, 1)) Then
            sAns = sAns & sChar
        End If
    Next
        
    OnlyAlphaNumericChars = sAns
End Function

Private Function IsAlphaNumeric(sChr As String) As Boolean
    IsAlphaNumeric = sChr Like "[0-9A-Za-z ]"
    'IsSafeChar = sChr Like "[,.!@$%^():;'""_+=0-9A-Za-z -]"
End Function

Public Function OnlySafeChars(ByVal OrigString As String) As String
    Dim lLen As Long
    Dim sAns As String
    Dim lCtr As Long
    Dim sChar As String
    
    OrigString = Trim(OrigString)
    lLen = Len(OrigString)
    For lCtr = 1 To lLen
        sChar = Mid(OrigString, lCtr, 1)
        If IsSafeChar(Mid(OrigString, lCtr, 1)) Then
            sAns = sAns & sChar
        End If
    Next
        
    OnlySafeChars = sAns
End Function

Private Function IsSafeChar(sChr As String) As Boolean
    IsSafeChar = sChr Like "[*0-9A-Za-z -]"
End Function

Public Function ScrubString(s As String) As String
    s = Replace(s, "&", "")
    s = Replace(s, "?", "")
    s = Replace(s, "%", "")
    s = Replace(s, "[", "")
    s = Replace(s, "]", "")
    s = Replace(s, "{", "")
    s = Replace(s, "}", "")
    s = Replace(s, "<", "")
    s = Replace(s, ">", "")
    s = Replace(s, "#", "")
    s = Replace(s, "(((", "~(~(~(")
    s = Replace(s, ")))", "~)~)~)")
    ScrubString = s
End Function

Public Function URLEncode(s As String, Optional SpaceAsPlus As Boolean = False) As String
    Dim StringLen As Long
    StringLen = Len(s)

    If s > 0 Then
        ReDim result(StringLen) As String
        Dim i As Long, CharCode As Integer
        Dim Char As String, Space As String

        If SpaceAsPlus Then Space = "+" Else Space = "%20"

        For i = 1 To StringLen
            Char = Mid$(s, i, 1)
            CharCode = Asc(Char)
            Select Case CharCode
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                    result(i) = Char
                Case 32
                    result(i) = Space
                Case 0 To 15
                    result(i) = "%0" & Hex(CharCode)
                Case Else
                    result(i) = "%" & Hex(CharCode)
            End Select
        Next i
        URLEncode = Join(result, "")
    End If
End Function

Public Function NormalizeSide(Side As String) As String
    Select Case "Side"
        Case "A"
            NormalizeSide = "A"
        Case "Aff"
            NormalizeSide = "A"
        Case "Pro"
            NormalizeSide = "A"
        Case "N"
            NormalizeSide = "N"
        Case "Neg"
            NormalizeSide = "N"
        Case "Con"
            NormalizeSide = "N"
        Case Else
            NormalizeSide = Side
    End Select
End Function

Public Function DisplaySide(Side As String, Optional EventName As String) As String
    If Side = "A" Or Side = "Aff" Or Side = "Pro" Then
        If EventName = "pf" Or EventName = "PF" Then
            DisplaySide = "Pro"
        Else
            DisplaySide = "Aff"
        End If
    End If

    If Side = "N" Or Side = "Neg" Or Side = "Con" Then
        If EventName = "pf" Or EventName = "PF" Then
            DisplaySide = "Con"
        Else
            DisplaySide = "Neg"
        End If
    End If
End Function

Public Function RoundName(Round As Variant) As String
    If IsNumeric(Round) Then
        RoundName = "Round " & Round
    Else
        RoundName = Round
    End If
End Function

Public Function HeadingToTitle(p As String) As String
    ' Clean text and ensure a non-zero string
    HeadingToTitle = Trim(OnlySafeChars(Replace(p, Chr(151), "-")))
    If Len(HeadingToTitle) > 1000 Then HeadingToTitle = Left(HeadingToTitle, 1000) 'Limit length to 1000 characters to avoid breaking XML
    If HeadingToTitle = "" Then HeadingToTitle = "-"
End Function

Public Function ConvertUnixTimestampToDate(ts As String) As Date
    ConvertUnixTimestampToDate = DateAdd("s", CDbl(ts), "1/1/1970")
End Function
