Attribute VB_Name = "Search"
'@IgnoreModule ProcedureNotUsed, ParameterNotUsed, EncapsulatePublicField
Option Explicit

Public SearchText As String

Public Sub SearchChanged(ByVal c As IRibbonControl, ByVal strText As Variant)
    ' Set the search text, then refresh ribbon
    SearchText = strText
    Ribbon.RefreshRibbon
    
    #If Mac Then
    #Else
        ' Activate the search box
        WordBasic.SendKeys "%d%s%r"
    #End If
End Sub

'@Ignore ProcedureCanBeWrittenAsFunction
Public Sub GetSearchResultsContent(ByVal c As IRibbonControl, ByRef returnedVal As Variant)
    Dim xml As String
    Dim SearchDir As String
    Dim x As Long
    
    On Error GoTo Handler
    
    ' Initialize XML
    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"
    
    ' If no search string, add a button with instructions
    If SearchText = "" Then
        xml = xml & "<button id=""SearchButton1"" label=""Press Enter to search."" />"
    Else
        #If Mac Then
            Dim Script As String
            Dim Raw As String
            Dim Results() As String
            Dim PathSegments() As String
            Dim r
            
            ' Construct SearchDir parameter - Use the UserProfile directory by default
            SearchDir = GetSetting("Verbatim", "Paperless", "SearchDir", "/Users/" & Environ("USER"))
            If SearchDir = "" Or SearchDir = "?" Then SearchDir = "/Users/" & Environ("USER")
            
            ' mdfind returns a list of results separated by newlines
            Script = "mdfind -onlyin '" & SearchDir & "' '" & SearchText & "'"
            Raw = AppleScriptTask("Verbatim.scpt", "RunShellScript", Script)
            If Len(Raw) = 0 Then
                xml = xml & "<button id=""SearchButton1"" label=""No results found."" />"
            Else
                Results = Split(Raw, Chr(13))
                x = 0
                For Each r In Results
                    x = x + 1
                    If x <= 25 And (LCase(Right(r, 4)) = "docx" Or LCase(Right(r, 3) = "doc")) Then
                        PathSegments = Split(r, Application.PathSeparator)
                        
                        xml = xml & "<button id=""SearchButton" & x & """ label=""" & PathSegments(UBound(PathSegments)) & """ "
                        xml = xml & "imageMso=""FileSaveAsWordDocx"" "
                        xml = xml & "tag=""" & Strings.URLEncode(CStr(r)) & """ "
                        xml = xml & "onAction=""Search.OpenSearchResult"" />"
                    End If
                Next r
            End If
        #Else
            Dim objConnection As Object
            Dim objRecordset As Object
            Dim SearchString As String
        
            ' Open ADO
            Set objConnection = CreateObject("ADODB.Connection")
            Set objRecordset = CreateObject("ADODB.Recordset")
            
            ' Construct SearchDir parameter - Use the UserProfile directory by default
            SearchDir = GetSetting("Verbatim", "Paperless", "SearchDir", CStr(Environ$("USERPROFILE")))
            If SearchDir = "" Then SearchDir = CStr(Environ$("USERPROFILE"))
            If Right$(SearchDir, 1) <> Application.PathSeparator Then SearchDir = SearchDir & Application.PathSeparator
            SearchString = "file:" & Replace(SearchDir, "\", "/")
            
            ' Set search string and open connection
            objConnection.Open "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"
            objRecordset.Open "SELECT Top 25 System.ItemName, System.ItemPathDisplay, System.ItemFolderPathDisplayNarrow, System.DateModified, System.Size FROM SystemIndex WHERE contains(System.Search.Contents, '""" & SearchText & """') and SCOPE='" & SearchString & "'", objConnection
    
            If objRecordset.EOF = True Then
                xml = xml & "<button id=""SearchButton1"" label=""No results found."" />"
            Else
                objRecordset.MoveFirst
                
                ' Loop returned records
                x = 0
                Do Until objRecordset.EOF
                    ' Add a button for top 25 returned document results
                    If x <= 25 And _
                        (Right$(objRecordset.Fields.Item("System.ItemName"), 4) = "docx" Or _
                        Right$(objRecordset.Fields.Item("System.ItemName"), 3) = "doc" Or _
                        Right$(objRecordset.Fields.Item("System.ItemName"), 3) = "rtf") Then
                        
                        xml = xml & "<button id=""SearchButton" & x & """ label=""" & objRecordset.Fields.Item("System.ItemName") & """ "
                        xml = xml & "supertip=""" _
                            & Strings.ScrubString(objRecordset.Fields.Item("System.ItemFolderPathDisplayNarrow")) _
                            & "&#10; &#10;" _
                            & "Date Modified:&#10;" _
                            & Strings.ScrubString(objRecordset.Fields.Item("System.DateModified")) _
                            & "&#10; &#10;" _
                            & "Size:&#10;" _
                            & Round(objRecordset.Fields.Item("System.Size") / 1024) _
                            & "KB" _
                            & "&#10; &#10;" _
                            & """ "
                    
                        xml = xml & "imageMso=""FileSaveAsWordDocx"" "
                        
                        xml = xml & "tag=""" & Strings.URLEncode(objRecordset.Fields.Item("System.ItemPathDisplay")) & """ "
                        xml = xml & "onAction=""Search.OpenSearchResult"" />"
                    End If
    
                    objRecordset.MoveNext
                    x = x + 1
                Loop
            End If
            
            ' Clean up ADO
            objRecordset.Close
            Set objRecordset = Nothing
            objConnection.Close
            Set objConnection = Nothing
        
            ' Add "more results" buttons
            xml = xml & "<button id=""ExplorerSearch"" label=""Search in Windows Explorer..."" supertip=""Opens your search in a Windows Explorer window"" imageMso=""NavigationPaneFind"" onAction=""Search.ExplorerSearch"" />"
            xml = xml & "<button id=""EverythingSearch"" label=""Search in Everything Search..."" supertip=""Opens your search in Everything Search (if installed)"" imageMso=""FindNext"" onAction=""Search.EverythingSearch"" />"
        #End If
    End If

    ' Close XML
    xml = xml & "</menu>"
    returnedVal = xml

    Exit Sub
    
Handler:
    #If Mac Then
    #Else
        Set objRecordset = Nothing
        Set objConnection = Nothing
    #End If
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Sub OpenSearchResult(ByVal c As IRibbonControl)
    ' Test for file extension, only open doc, docx, rtf - otherwise call shell
    If Right$(c.Tag, 4) = "docx" Or Right$(c.Tag, 3) = "doc" Or Right$(c.Tag, 3) = "rtf" Then
        Documents.Open Strings.URLDecode(c.Tag)
    Else
        #If Mac Then
            MsgBox "Word can't open this kind of file!"
        #Else
            CreateObject("WScript.Shell").Open Strings.URLDecode(c.Tag)
        #End If
    End If
End Sub

Public Sub ExplorerSearch(ByVal c As IRibbonControl)
    #If Mac Then
        MsgBox "Explorer Search is not supported on Mac."
        Exit Sub
    #Else
        Dim SearchDir As String
        
        ' Use user profile as the default search location if not set
        SearchDir = GetSetting("Verbatim", "Paperless", "SearchDir", CStr(Environ$("USERPROFILE")))
        If SearchDir = "" Then SearchDir = CStr(Environ$("USERPROFILE"))
        If Right$(SearchDir, 1) <> Application.PathSeparator Then SearchDir = SearchDir & Application.PathSeparator
        
        Shell "explorer ""search-ms://query=" & SearchText & "&crumb=location:" & SearchDir & """", vbNormalFocus
    #End If
End Sub

Public Sub EverythingSearch(ByVal c As IRibbonControl)
    #If Mac Then
        MsgBox "Everything Search is not supported on Mac."
        Exit Sub
    #Else
        Dim SearchDir As String
        
        ' Use user profile as the default search location if not set
        SearchDir = GetSetting("Verbatim", "Paperless", "SearchDir", CStr(Environ$("USERPROFILE")))
        If SearchDir = "" Then SearchDir = CStr(Environ$("USERPROFILE"))
        If Right$(SearchDir, 1) <> Application.PathSeparator Then SearchDir = SearchDir & Application.PathSeparator
        
        Dim SearchPath As String
        SearchPath = GetSetting("Verbatim", "Plugins", "SearchPath", "")
        If SearchPath <> "" Then
            If Filesystem.FileExists(SearchPath) = False Then
                MsgBox "External Search program not found. Please check the path to the application in your Verbatim settings, or remove it to use the Everything Search plugin."
                Exit Sub
            Else
                CreateObject("WSCript.Shell").Run SearchPath, 0, True
                Exit Sub
            End If
        End If
        
        Dim EverythingPath As String
        
        If Filesystem.FileExists(Environ$("ProgramW6432") _
            & Application.PathSeparator _
            & "Verbatim" _
            & Application.PathSeparator _
            & "Plugins" _
            & Application.PathSeparator _
            & "Search" _
            & Application.PathSeparator _
            & "Everything.exe" _
        ) = True Then
            EverythingPath = Environ$("ProgramW6432") _
                & Application.PathSeparator _
                & "Verbatim" _
                & Application.PathSeparator _
                & "Plugins" _
                & Application.PathSeparator _
                & "Search" _
                & Application.PathSeparator _
                & "Everything.exe"
        ElseIf Filesystem.FileExists(Environ$("ProgramW6432") & Application.PathSeparator & "Everything" & Application.PathSeparator & "Everything.exe") = True Then
            EverythingPath = Environ$("ProgramW6432") & Application.PathSeparator & "Everything" & Application.PathSeparator & "Everything.exe"
        Else
            MsgBox "Everything Search plugin must be installed to use this feature. Please see https://paperlessdebate.com/ for details on how to install."
            Exit Sub
        End If
        
        Shell EverythingPath & " -s """ & SearchDir & " " & SearchText & """", vbNormalFocus
    #End If
End Sub
