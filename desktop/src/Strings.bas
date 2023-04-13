Attribute VB_Name = "Strings"
Option Explicit

Public Function OnlyAlphaNumericChars(ByVal OrigString As String) As String
    Dim lLen As Long
    Dim sAns As String
    Dim lCtr As Long
    Dim sChar As String
    Dim TrimString As String
    
    TrimString = Trim$(OrigString)
    lLen = Len(TrimString)
    For lCtr = 1 To lLen
        sChar = Mid$(TrimString, lCtr, 1)
        If IsAlphaNumeric(Mid$(TrimString, lCtr, 1)) Then
            sAns = sAns & sChar
        End If
    Next
        
    OnlyAlphaNumericChars = sAns
End Function

Public Function IsAlphaNumeric(ByVal sChr As String) As Boolean
    IsAlphaNumeric = sChr Like "[0-9A-Za-z ]"
End Function

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
    'IsSafeChar = sChr Like "[,.!@$%^():;'""_+=0-9A-Za-z -]"
End Function

Public Function ScrubString(ByRef s As String) As String
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

Public Function URLEncode(ByRef s As String, Optional ByVal SpaceAsPlus As Boolean = False) As String
    Dim StringLen As Long
    StringLen = Len(s)

    If StringLen > 0 Then
        ReDim Result(StringLen) As String
        Dim i As Long
        Dim CharCode As Long
        Dim Char As String
        Dim Space As String

        If SpaceAsPlus Then Space = "+" Else Space = "%20"

        For i = 1 To StringLen
            Char = Mid$(s, i, 1)
            CharCode = Asc(Char)
            Select Case CharCode
                Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
                    Result(i) = Char
                Case 32
                    Result(i) = Space
                Case 0 To 15
                    Result(i) = "%0" & Hex$(CharCode)
                Case Else
                    Result(i) = "%" & Hex$(CharCode)
            End Select
        Next i
        URLEncode = Join(Result, "")
    End If
End Function

Public Function URLDecode(ByVal s As String) As String
    On Error GoTo Handler
    
    Dim i As Long
    Dim retval As String
    Dim tmp As String
    
    If Len(s) > 0 Then
        ' Loop through each char
        For i = 1 To Len(s)
            tmp = Mid$(s, i, 1)
            tmp = Replace(tmp, "+", " ")
            If tmp = "%" And Len(s) + 1 > i + 2 Then
                tmp = Mid$(s, i + 1, 2)
                tmp = Chr$(CDbl("&H" & tmp))
                i = i + 2
            End If
            retval = retval & tmp
        Next
        URLDecode = retval
    End If

    Exit Function

Handler:
    URLDecode = ""
End Function

Public Function NormalizeSide(ByVal Side As String) As String
    Select Case "Side"
        Case Is = "A"
            NormalizeSide = "A"
        Case Is = "Aff"
            NormalizeSide = "A"
        Case Is = "Pro"
            NormalizeSide = "A"
        Case Is = "N"
            NormalizeSide = "N"
        Case Is = "Neg"
            NormalizeSide = "N"
        Case Is = "Con"
            NormalizeSide = "N"
        Case Else
            NormalizeSide = Side
    End Select
End Function

Public Function DisplaySide(ByVal Side As String, Optional ByVal EventName As String) As String
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

Public Function RoundName(ByRef Round As Variant) As String
    If IsNumeric(Round) Then
        RoundName = "Round " & Round
    Else
        RoundName = Round
    End If
End Function

Public Function HeadingToTitle(ByVal p As String) As String
    ' Clean text and ensure a non-zero string
    HeadingToTitle = Trim$(OnlySafeChars(Replace(p, Chr$(151), "-")))
    If Len(HeadingToTitle) > 1000 Then HeadingToTitle = Left$(HeadingToTitle, 1000) 'Limit length to 1000 characters to avoid breaking XML
    If HeadingToTitle = "" Then HeadingToTitle = "-"
End Function

'@Ignore ProcedureNotUsed
Public Function ConvertUnixTimestampToDate(ByRef ts As String) As Date
    ConvertUnixTimestampToDate = DateAdd("s", CDbl(ts), "1/1/1970")
End Function
