Attribute VB_Name = "Search"
Option Explicit

Public SearchText As String

Sub SearchChanged(c As IRibbonControl, strText As String)
    ' Set the search text, then refresh ribbon
    SearchText = strText
    Ribbon.RefreshRibbon
    
    ' Activate the search box
    WordBasic.SendKeys "%d%s%r"
    
End Sub

Sub GetSearchResultsContent(c As IRibbonControl, ByRef returnedVal)
    #If Mac Then
        ' TODO - figure out a Mac version
        MsgBox "Search is not supported on Mac."
        Exit Sub
    #Else
        Dim objConnection As Object
        Dim objRecordset As Object
        
        Dim x As Long
        Dim xml As String
        Dim SearchDir As String
        
        On Error GoTo Handler
        
        ' Initialize XML
        xml = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"
        
        ' If no search string, add a button with instructions
        If SearchText = "" Then
            xml = xml & "<button id=""SearchButton1"" label=""Press Enter to search."" />"
        Else
        
            ' Open ADO
            Set objConnection = CreateObject("ADODB.Connection")
            Set objRecordset = CreateObject("ADODB.Recordset")
            
            ' Construct SearchDir parameter - Use the UserProfile directory by default
            SearchDir = GetSetting("Verbatim", "Paperless", "SearchDir", CStr(Environ("USERPROFILE")))
            If SearchDir = "" Then SearchDir = CStr(Environ("USERPROFILE"))
            If Right(SearchDir, 1) <> Application.PathSeparator Then SearchDir = SearchDir & Application.PathSeparator
            SearchDir = "file:" & Replace(GetSetting("Verbatim", "Paperless", "SearchDir", CStr(Environ("USERPROFILE"))), "\", "/")
            
            ' Set search string and open connection
            objConnection.Open "Provider=Search.CollatorDSO;Extended Properties='Application=Windows';"
            objRecordset.Open "SELECT Top 25 System.ItemName, System.ItemPathDisplay, System.ItemFolderPathDisplayNarrow, System.DateModified, System.Size FROM SystemIndex WHERE contains(System.Search.Contents, '""" & SearchText & """') and SCOPE='" & SearchDir & "'", objConnection
    
            If objRecordset.EOF = True Then
                xml = xml & "<button id=""SearchButton1"" label=""No results found."" />"
            Else
                objRecordset.MoveFirst
                
                ' Loop returned records
                x = 0
                Do Until objRecordset.EOF
                    ' Add a button for each returned result
                    xml = xml & "<button id=""SearchButton" & x & """ label=""" & objRecordset.Fields.Item("System.ItemName") & """ "
                    xml = xml & "supertip=""" & objRecordset.Fields.Item("System.ItemFolderPathDisplayNarrow") & "&#10; &#10;" & _
                    "Date Modified:&#10;" & objRecordset.Fields.Item("System.DateModified") & "&#10; &#10;" & _
                    "Size:&#10;" & Round(objRecordset.Fields.Item("System.Size") / 1024) & "KB" & "&#10; &#10;" & """ "
                    If Right(objRecordset.Fields.Item("System.ItemName"), 4) = "docx" Or _
                        Right(objRecordset.Fields.Item("System.ItemName"), 3) = "doc" Or _
                        Right(objRecordset.Fields.Item("System.ItemName"), 3) = "rtf" Then
                            xml = xml & "imageMso=""FileSaveAsWordDocx"" "
                    End If
                    
                    xml = xml & "tag=""" & objRecordset.Fields.Item("System.ItemPathDisplay") & """ "
                    xml = xml & "onAction=""Search.OpenSearchResult"" />"
    
                    objRecordset.MoveNext
                    x = x + 1
                Loop
            End If
            
            ' Clean up ADO
            objRecordset.Close
            Set objRecordset = Nothing
            objConnection.Close
            Set objConnection = Nothing
        
            ' Add a "more results" button
            xml = xml & "<button id=""MoreResults"" label=""More results..."" supertip=""Opens your search in an explorer window"" onAction=""Search.ExplorerSearch"" />"
            
        End If
        
        ' Close XML
        xml = xml & "</menu>"
        
        Debug.Print xml
        returnedVal = xml
    #End If
    Exit Sub
    
Handler:
    Set objRecordset = Nothing
    Set objConnection = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub OpenSearchResult(c As IRibbonControl)
    #If Mac Then
        ' TODO - figure out a Mac version
        MsgBox "Search is not supported on Mac."
        Exit Sub
    #Else
        Dim s As Object
        Set s = CreateObject("WScript.Shell")
        
        ' Test for file extension, only open doc, docx, rtf - otherwise call shell
        If Right(c.Tag, 4) = "docx" Or Right(c.Tag, 3) = "doc" Or Right(c.Tag, 3) = "rtf" Then
            Documents.Open c.Tag
        Else
            Set s = CreateObject("WScript.Shell")
            s.Open (c.Tag)
            Set s = Nothing
        End If
    #End If
End Sub

Sub ExplorerSearch(c As IRibbonControl)
    #If Mac Then
        ' TODO - figure out a Mac version
        MsgBox "Search is not supported on Mac."
        Exit Sub
    #Else
        Dim SearchDir As String
        
        ' Construct SearchDir, then pass it to the shell
        SearchDir = GetSetting("Verbatim", "Paperless", "SearchDir", CStr(Environ("USERPROFILE")))
        If SearchDir = "" Then SearchDir = CStr(Environ("USERPROFILE"))
        If Right(SearchDir, 1) <> Application.PathSeparator Then SearchDir = SearchDir & Application.PathSeparator
        
        Shell "explorer ""search-ms://query=" & SearchText & "&crumb=location:" & SearchDir & """", vbNormalFocus
    #End If
End Sub
