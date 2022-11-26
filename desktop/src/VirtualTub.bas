Attribute VB_Name = "VirtualTub"
Option Explicit

Sub GetVTubContent(control As IRibbonControl, ByRef returnedVal)
'Get content for dynamic menu from XML file

    ' TODO - rewrite to use JSON -> XML

    Dim VTubPath As String
    Dim VTubFolder As Folder
    Dim FileNumber As Integer
    
    ' Skip Errors
    On Error Resume Next
        
    ' Get VTubPath from Settings and make sure it exists
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath", "")
    If VTubPath = "" Or VTubPath = "?" Then
        If MsgBox("You haven't configured a VTub location in the Verbatim settings. Open Settings?", vbYesNo) = vbYes Then
            UI.ShowForm "Settings"
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    ' Append trailing \ if missing
    If Right(VTubPath, 1) <> Application.PathSeparator Then VTubPath = VTubPath & Application.PathSeparator
    
    ' Check if VTubXMLFile exists
    If Filesystem.FileExists(VTubPath & "VTub.xml") = False Then
        ' If no XML file, return a button to create it
        returnedVal = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"
        returnedVal = returnedVal & "<button id=""CreateVTub"" label=""Create VTub"" onAction=""VirtualTub.VTubCreateButton"" imageMso=""_3DSurfaceMaterialClassic""" & " />"
        returnedVal = returnedVal & "</menu>"
        Exit Sub
    End If

    ' If VTubRefreshPrompt is turned on, check if Tub is out of date by comparing date modified of files to XML file
    ' Have to loop all files becuase Word changes folder modified date when opening docs
    If GetSetting("Verbatim", "VTub", "VTubRefreshPrompt", True) = True Then
        Set VTubFolder = Filesystem.GetFolder(VTubPath)
        VTubDepth = 0
        VTubMaxDepth = 0
        VTubLastModified = ""
        Call VirtualTub.VTubFileCounterRecursion(VTubFolder)
        If Filesystem.GetFile(VTubXMLFileName).DateLastModified < VTubLastModified Then
            If MsgBox("The VTub has not been refreshed since you last changed files. Refresh Now?", vbYesNo) = vbYes Then
                Call VirtualTub.VTubRefresh
                Set VTubFolder = Nothing
                Exit Sub
            End If
        End If
    End If

    ' Open and read the XML file
    FileNumber = FreeFile
    Open VTubXMLFileName For Input As #FileNumber
    returnedVal = Input$(LOF(FileNumber), FileNumber)
    Close #FileNumber
    
    
    
    Set VTubFolder = Nothing
    Exit Sub

Handler:
    Set VTubFolder = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub VTubRefreshButton(control As IRibbonControl)
    VirtualTub.VTubRefresh
End Sub
Sub VTubCreateButton(control As IRibbonControl)
    VirtualTub.VTubCreate
End Sub
Sub VTubSettingsButton(control As IRibbonControl)
    UI.ShowForm "Settings"
End Sub

Sub VTubInsertBookmark(control As IRibbonControl)
    ' Insert bookmark - get the file path and bookmark name by splitting the tag attribute on the !#! delimiter
    Selection.InsertFile Split(control.Tag, "!#!", 2)(0), Split(control.Tag, "!#!", 2)(1)
End Sub

Private Sub VTubCreate()
    Dim Folder As clsFolder
    Dim Subfolder As clsFolder
    Dim File As clsFile
    
    Dim RootMenu As Dictionary
    Dim SubfolderMenu As Dictionary
    Dim FileMenu As Dictionary
    Dim Headings As Dictionary
    Dim Children As Dictionary
        
    Dim FileCount As Long
    FileCount = 0
    
    Dim DepthExceeded As Boolean
    DepthExceeded = False
    
    Dim i As Long
    Dim j As Long
    
    Dim JSON
    
    On Error GoTo Handler
    
    Set RootMenu = New Dictionary
    
    Dim VTubPath As String
     
    ' Get VTubPath from settings
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath", "")
    
    If VTubPath = vbNullString Then
        MsgBox "You must configure a folder for the VTub first"
        If MsgBox("You haven't configured a folder for the VTub. Open Settings?", vbYesNo, "Open Settings?") = vbYes Then
            UI.ShowForm "Settings"
        End If
        Exit Sub
    End If
    
    ' Append trailing \ if missing
    If Right(VTubPath, 1) <> Application.PathSeparator Then VTubPath = VTubPath & Application.PathSeparator
    
    Set Folder = GetFolder(VTubPath)
      
    For i = 1 To Folder.Subfolders.Count
        Set Subfolder = GetFolder(Folder.Subfolders(i))
        If Subfolder.Subfolders.Count > 0 Then DepthExceeded = True
        FileCount = FileCount + Subfolder.Files.Count
    Next
    
    FileCount = FileCount + Folder.Files.Count
           
    If FileCount > 20 Then
        If MsgBox("You have a large number of files (>20) in the VTub. This could take a few minutes - okay?", vbYesNo, "You sure?") = vbNo Then Exit Sub
    End If
    
    If DepthExceeded = True Then MsgBox "VTub can only handle one level of subfolders - files deeper than one subfolder will be ignored.", vbOKOnly

    ' Show progress bar
    Dim ProgressForm As frmProgress
    Set ProgressForm = New frmProgress
    ProgressForm.Caption = "Creating VTub..."
    ProgressForm.lblProgress.Width = 0
    ProgressForm.lblCaption.Caption = "File 0 of " & FileCount
    ProgressForm.Show

    RootMenu.Add "FileCount", FileCount
    RootMenu.Add "FolderCount", Folder.Subfolders.Count

   ' Iterate through each subfolder in first depth level - XML menu is limited to 5 levels and we need 3 for the file contents
    For i = 1 To Folder.Subfolders.Count
       ' Trap for cancel button on Progress Form
        If ProgressForm.Visible = False Then Exit Sub
        
        ' TODO - need to use a filecounter outside the loop to capture both outer and inner loop files
        Dim ProgressPct
        ProgressPct = i / FileCount
        ProgressForm.lblCaption.Caption = Str(Round(ProgressPct * 100, 0)) & "% - " & "Processing File " & i & " of " & FileCount
        ProgressForm.lblFile.Caption = "Processing " & f.Name
        ProgressForm.lblProgress.Width = ProgressPct * ProgressForm.fProgress.Width
        If ProgressForm.lblProgress.Width > 15 Then ProgressForm.lblProgress.Width = ProgressForm.lblProgress.Width - 15
        
        DoEvents ' Necessary for Progress form to update
    
        Set Subfolder = GetFolder(Folder.Subfolders(i))
        Set SubfolderMenu = New Dictionary
        SubfolderMenu.Add "MenuType", "Folder"
        SubfolderMenu.Add "Name", Subfolder.Name
        SubfolderMenu.Add "Path", Subfolder.Path
        
        Set Children = New Dictionary
        
        For j = 1 To Subfolder.Files.Count
            If Right(Subfolder.Files(j), 4) = "docx" And Left(Subfolder.Files(j), 1) <> "~" Then
                Set File = GetFile(Subfolder.Files(j))
                Set FileMenu = New Dictionary
                FileMenu.Add "MenuType", "File"
                FileMenu.Add "Name", File.Name
                FileMenu.Add "Path", File.Path
                
                Set Headings = VirtualTub.AddBookmarks(File.Path)
                
                FileMenu.Add "Children", Headings
                Set File = GetFile(Subfolder.Files(j))
                FileMenu.Add "DateLastModified", Format(File.DateLastModified)
                
                Children.Add Replace(File.Path, "\", "\\"), FileMenu
            End If
        Next
        
        SubfolderMenu.Add "Children", Children
        SubfolderMenu.Add "DateLastModified", Format(Subfolder.DateLastModified)

        RootMenu.Add Replace(Subfolder.Path, "\", "\\"), SubfolderMenu
    Next
    
    For i = 1 To Folder.Files.Count
        If Right(Folder.Files(i), 4) = "docx" Then
            Set File = GetFile(Folder.Files(i))
            Set FileMenu = New Dictionary
            FileMenu.Add "MenuType", "File"
            FileMenu.Add "Name", File.Name
            FileMenu.Add "Path", File.Path
            
            Set Headings = VirtualTub.AddBookmarks(File.Path)
            
            FileMenu.Add "Children", Headings
            Set File = GetFile(Folder.Files(i))
            FileMenu.Add "DateLastModified", Format(File.DateLastModified)
            
            RootMenu.Add Replace(File.Path, "\", "\\"), FileMenu
        End If
    Next
    
    Set Folder = GetFolder(GetSetting("Verbatim", "VTub", "VTubPath", ""))
    RootMenu.Add "DateLastModified", Format(Folder.DateLastModified)
        
    JSON = JSONTools.ConvertToJson(RootMenu)
    Debug.Print JSON
       
    ' Save file
    Dim VTubFilePath
    Dim OutputFile
    VTubFilePath = GetSetting("Verbatim", "VTub", "VTubPath", vbNullString) & Application.PathSeparator & "VTub.json"
    OutputFile = FreeFile
    Open VTubFilePath For Output As #OutputFile
    Print #OutputFile, JSON
    Close #OutputFile
    
    ' Update progress form as complete
    ProgressForm.lblCaption.Caption = "Processing complete."
    ProgressForm.lblFile.Caption = ""
    ProgressForm.lblProgress.Width = ProgressForm.fProgress.Width - 6
    Unload ProgressForm
    Set ProgressForm = Nothing
    
    Ribbon.RefreshRibbon
    MsgBox "VTub successfully created!" & vbCrLf & vbCrLf & "If you get an error when you click OK that ""The document is too large to save. Delete some text before saving."", you can ignore it - it's a bug in Word and won't affect the VTub."
    
    Exit Sub
    
Handler:
    Unload ProgressForm
    Set ProgressForm = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Sub

Sub RefreshVTub()
    Dim VTubPath
    Dim JSON As String
    Dim RootMenu As Object
    Dim Menu As Object
    Dim Children As Object
    Dim xml As String
    
    On Error GoTo Handler
    
    ' Verify before proceeding
    If MsgBox("Are you sure you want to refresh the VTub?", vbOKCancel) = vbCancel Then Exit Sub
    
    ' Get VTubPath from Settings
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath")
    If Right(VTubPath, 1) <> Application.PathSeparator Then VTubPath = VTubPath & Application.PathSeparator
    
    ' Load JSON from file
    JSON = Filesystem.ReadFile(VTubPath & "VTub.json")

    Set RootMenu = JSONTools.ParseJson(JSON)
    
    Dim Folder As clsFolder
    Dim FileCount As Long
    FileCount = 0
       
    Set Folder = GetFolder(GetSetting("Verbatim", "VTub", "VTubPath", ""))
    
    Dim i As Long
    Dim Subfolder
    For i = 1 To Folder.Subfolders.Count
        Set Subfolder = GetFolder(Folder.Subfolders(i))
        FileCount = FileCount + Subfolder.Files.Count
    Next
    
    FileCount = FileCount + Folder.Files.Count
    
    If (CInt(FileCount) <> CInt(RootMenu("FileCount")) Or CInt(Folder.Subfolders.Count) <> CInt(RootMenu("FolderCount"))) Then
        If MsgBox("The number of files or folders in your VTub appear to have changed and needs to be rebuilt from scratch. Rebuild now?", vbOKCancel) = vbCancel Then Exit Sub
        VirtualTub.TestVTub
        Exit Sub
    End If
    
    ' Show progress bar
    Dim ProgressForm As frmProgress
    Set ProgressForm = New frmProgress
    ProgressForm.Caption = "Refreshing VTub..."
    ProgressForm.lblProgress.Width = 0
    ProgressForm.lblCaption.Caption = "File 0 of " & FileCount
    ProgressForm.Show
    
    Dim key As Variant
    Dim subkey As Variant
    Dim Child
    Dim File
    For Each key In RootMenu.Keys
    
       ' Trap for cancel button on Progress Form
        If ProgressForm.Visible = False Then Exit Sub
        
        ' TODO - need to use a filecounter outside the loop to capture both outer and inner loop files
        Dim ProgressPct
        ProgressPct = i / FileCount
        ProgressForm.lblCaption.Caption = Str(Round(ProgressPct * 100, 0)) & "% - " & "Processing File " & i & " of " & FileCount
        ProgressForm.lblFile.Caption = "Processing " & f.Name
        ProgressForm.lblProgress.Width = ProgressPct * ProgressForm.fProgress.Width
        If ProgressForm.lblProgress.Width > 15 Then ProgressForm.lblProgress.Width = ProgressForm.lblProgress.Width - 15
        
        DoEvents ' Necessary for Progress form to update

        If key <> "FileCount" And key <> "FolderCount" And key <> "DateLastModified" Then
            Set Menu = RootMenu(key)
            If (Menu("MenuType") = "Folder" And Menu.Exists("Children")) Then
                Set Children = Menu("Children")
    
                For Each subkey In Children.Keys
                    Set Child = Children(subkey)
                    If Child("MenuType") = "File" Then
                        Dim Path As String
                        Path = Child("Path")
                        Set File = GetFile(Path)
                        If Child("DateLastModified") <> Format(File.DateLastModified) Then
                            Set Child("Children") = VirtualTub.AddBookmarks(File.Path)
                            Set File = GetFile(File.Path)
                            Child("DateLastModified") = Format(File.DateLastModified)
                            'Set Children(Replace(subkey, "\", "\\")) = Child
                        End If
                    End If
                    
                    Children.Remove subkey
                    Children.Add Replace(subkey, "\", "\\"), Child
                Next subkey
                
                Path = Menu("Path")
                Set Folder = GetFolder(Path)
                Menu("DateLastModified") = Format(Folder.DateLastModified)
                'Set RootMenu(Replace(key, "\", "\\")) = Menu
            ElseIf Menu("MenuType") = "File" Then
                Path = Menu("Path")
                Set File = GetFile(Path)
                If Menu("DateLastModified") <> Format(File.DateLastModified) Then
                    Set Menu("Children") = VirtualTub.AddBookmarks(File.Path)
                    Set File = GetFile(File.Path)
                    Menu("DateLastModified") = Format(File.DateLastModified)
                    'Set RootMenu(Replace(key, "\", "\\")) = Menu
                End If
            End If
            
            RootMenu.Remove key
            RootMenu.Add Replace(key, "\", "\\"), Menu
        End If
    Next key
    
    Set Folder = GetFolder(GetSetting("Verbatim", "VTub", "VTubPath", vbNullString))
    RootMenu("DateLastModified") = Format(Folder.DateLastModified)
  
    JSON = JSONTools.ConvertToJson(RootMenu)
  
    ' Update progress form as complete
    ProgressForm.lblCaption.Caption = "Processing complete."
    ProgressForm.lblFile.Caption = ""
    ProgressForm.lblProgress.Width = ProgressForm.fProgress.Width - 6
  
    ' Save file
    Dim VTubFilePath
    Dim OutputFile
    VTubFilePath = GetSetting("Verbatim", "VTub", "VTubPath", vbNullString) & Application.PathSeparator & "VTub.json"
    OutputFile = FreeFile
    Open VTubFilePath For Output As #OutputFile
    Print #OutputFile, JSON
    Close #OutputFile
    
    Unload ProgressForm
    Set ProgressForm = Nothing
    
    ' Refresh ribbon and notify
    Ribbon.RefreshRibbon
    MsgBox "VTub successfully refreshed!" & vbCrLf & vbCrLf & "If you get an error when you click OK that ""The document is too large to save. Delete some text before saving."", you can ignore it - it's a bug in Word and won't affect the VTub."
    
    Exit Sub
    
Handler:
    Unload ProgressForm
    Set ProgressForm = Nothing
    
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Function ConvertDictionaryToXML(d) As String
    Dim xml As String
    Dim Children
    Dim Child
    Dim key As Variant
    
    VTubMenuIDNumber = VTubMenuIDNumber + 1
     
    If d.Exists("Children") Then
        Set Children = d("Children")
        
        If d("MenuType") = "Heading" Then
            xml = xml & "<splitButton "
            xml = xml & "id=""VTub" & VTubMenuIDNumber & """>" & vbCrLf
       
            VTubMenuIDNumber = VTubMenuIDNumber + 1
            xml = xml & "<button "
            xml = xml & "id=""VTub" & VTubMenuIDNumber & """ "
            xml = xml & "label=""" & d("Label") & """ "
            xml = xml & "tag=""" & d("Path") & "!#!" & d("Name") & """ "
            xml = xml & "onAction=""VirtualTub.VTubInsertBookmark"" "
            xml = xml & "imageMso=""ExportTextFile"" "
            xml = xml & "/>" & vbCrLf
           
            VTubMenuIDNumber = VTubMenuIDNumber + 1
            xml = xml & "<menu "
            xml = xml & "id=""VTub" & VTubMenuIDNumber & """ "
            xml = xml & ">" & vbCrLf
        Else
            xml = xml & "<menu "
            xml = xml & "id=""VTub" & VTubMenuIDNumber & """ "
            xml = xml & "label=""" & d("Name") & """ "
            If d("MenuType") = "Folder" Then
                xml = xml & "imageMso=""Folder"" "
            Else
                xml = xml & "imageMso=""FileSaveAsWordDocx"" "
            End If
            xml = xml & "tag=""" & d("Path") & "!#!" & d("DateLastModified") & """"
            xml = xml & ">" & vbCrLf
        End If
        
        For Each key In Children.Keys
            Set Child = Children(key)
            If Child.Exists("Children") Then
                xml = xml & ConvertDictionaryToXML(Child)
            Else
                VTubMenuIDNumber = VTubMenuIDNumber + 1
                xml = xml & "<button "
                xml = xml & "id=""VTub" & VTubMenuIDNumber & """ "
                xml = xml & "label=""" & Child("Label") & """ "
                xml = xml & "tag=""" & Child("Path") & "!#!" & Child("Name") & """ "
                xml = xml & "onAction=""VirtualTub.VTubInsertBookmark"" "
                xml = xml & "imageMso=""ExportTextFile"" "
                xml = xml & "/>" & vbCrLf
            End If
        Next key
        
        xml = xml & "</menu>" & vbCrLf
        If d("MenuType") = "Heading" Then
            xml = xml & "</splitButton>" & vbCrLf
        End If
    Else
        If d("MenuType") = "Heading" Then
            xml = xml & "<button "
            xml = xml & "id=""VTub" & VTubMenuIDNumber & """ "
            xml = xml & "label=""" & d("Label") & """ "
            xml = xml & "tag=""" & d("Path") & "!#!" & d("Name") & """ "
            xml = xml & "onAction=""VirtualTub.VTubInsertBookmark"" "
            xml = xml & "imageMso=""ExportTextFile"" "
            xml = xml & "/>" & vbCrLf
        Else
            xml = xml & "<menu "
            xml = xml & "id=""VTub" & VTubMenuIDNumber & """ "
            xml = xml & "label=""" & d("Name") & """ "
            If d("MenuType") = "Folder" Then
                xml = xml & "imageMso=""Folder"" "
            Else
                xml = xml & "imageMso=""FileSaveAsWordDocx"" "
            End If
            xml = xml & "tag=""" & d("Path") & "!#!" & d("DateLastModified") & """"
            xml = xml & ">" & vbCrLf
            xml = xml & "</menu>" & vbCrLf
        End If
    End If
    
    ConvertDictionaryToXML = xml
End Function

Sub ConvertVTubToXML()
    Dim VTubPath
    Dim JSON As String
    Dim RootMenu As Object
    Dim Menu
    Dim Children As Object
    Dim xml As String
    
    On Error GoTo Handler
    
    ' Get VTubPath from Settings
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath")
    If Right(VTubPath, 1) <> Application.PathSeparator Then VTubPath = VTubPath & Application.PathSeparator
    
    ' Load JSON from file
    JSON = Filesystem.ReadFile(VTubPath & "VTub.json")
    ' Debug.Print JSON

    Set RootMenu = JSONTools.ParseJson(JSON)

    xml = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">" & vbCrLf
    
    Dim key As Variant
    For Each key In RootMenu.Keys
        If key <> "FileCount" And key <> "FolderCount" And key <> "DateLastModified" Then
            Set Menu = RootMenu(key)
            xml = xml & ConvertDictionaryToXML(Menu)
        End If
    Next key
  
    ' Add default buttons
    xml = xml & "<menuSeparator id=""VTubSeparator"" />"
    xml = xml * "<button id=""RefreshVTub"" label=""Refresh VTub"" onAction=""VirtualTub.VTubRefreshButton"" imageMSO=""AccessRefreshAllLists"" />"
    xml = xml * "<button id=""RecreateVTub"" label=""Recreate VTub"" onAction=""VirtualTub.VTubCreateButton"" imageMSO=""_3DSurfaceMaterialClassic"" />"
    xml = xml * "<button id=""VTubSettings"" label=""VTub Settings"" onAction=""VirtualTub.VTubSettingsButton"" imageMSO=""_3DLightingFlatClassic"" />"
    
    xml = xml & "</menu>"
    
    Debug.Print xml
    
    'Save file
    Dim OutputFile
    Dim VTubFilePath
    VTubFilePath = "C:\Users\hardy\Desktop\Tub\VTub.xml"
    OutputFile = FreeFile
    Open VTubFilePath For Output As #OutputFile
    Print #OutputFile, xml
    Close #OutputFile
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

Public Function HeadingTitle(p As String) As String
    ' Clean text and ensure a non-zero string
    HeadingTitle = Trim(OnlySafeChars(Replace(p, Chr(151), "-")))
    If Len(HeadingTitle) > 1000 Then HeadingTitle = Left(HeadingTitle, 1000) 'Limit length to 1000 characters to avoid breaking XML
    If HeadingTitle = "" Then HeadingTitle = "-"
End Function

Function AddBookmarks(Path As String) As Dictionary
    Dim pCount As Long
    Dim p As Paragraph
    Dim pp As Paragraph
    Dim ppp As Paragraph
      
    On Error GoTo Handler
      
    ' Open the file in the background and activate it
    Documents.Open FileName:=Path, Visible:=False
    Documents(Path).Activate
    
    ' Delete all bookmarks
    VirtualTub.RemoveBookmarks
    
    Dim Bookmarks As Dictionary
    Set Bookmarks = New Dictionary
    Dim Level1Menu As Dictionary
    Dim Level1Children As Dictionary
    Dim Level2Menu As Dictionary
    Dim Level2Children As Dictionary
    Dim Level3Menu As Dictionary
    
    pCount = 0
    
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
    
    Dim StartHeading As Integer
    StartHeading = VirtualTub.LargestHeading
    
    Selection.HomeKey Unit:=wdStory
    Selection.Collapse
    
    Dim SubHeadingLevel As Integer
    Dim Level1Range As Range
    Dim Level2Range As Range
    Dim Level3Range As Range
    
    For Each p In Documents(Path).Paragraphs
        pCount = pCount + 1
        If p.outlineLevel = StartHeading Then
            Set Level1Range = Paperless.SelectHeadingAndContentRange(p)
            Documents(Path).Bookmarks.Add "Level_1_" & pCount, Level1Range
            Set Level1Menu = New Dictionary
            Set Level1Children = New Dictionary
            Level1Menu.Add "Name", "Level_1_" & pCount
            Level1Menu.Add "Path", Path
            Level1Menu.Add "Label", HeadingTitle(p.Range.Text)
            Level1Menu.Add "MenuType", "Heading"
            Level1Menu.Add "Children", Level1Children
            Bookmarks.Add "Level_1_" & pCount, Level1Menu
            
            SubHeadingLevel = 3
            For Each pp In Level1Range.Paragraphs
                If pp.outlineLevel = wdOutlineLevel2 Then SubHeadingLevel = 2
            Next pp
            
            For Each pp In Level1Range.Paragraphs
                pCount = pCount + 1
                If pp.outlineLevel = SubHeadingLevel And pp.outlineLevel > StartHeading Then
                    Set Level2Range = Paperless.SelectHeadingAndContentRange(pp)
                    Documents(Path).Bookmarks.Add "Level_2_" & pCount, Level2Range
                    Set Level2Menu = New Dictionary
                    Set Level2Children = New Dictionary
                    Level2Menu.Add "Name", "Level_2_" & pCount
                    Level2Menu.Add "Path", Path
                    Level2Menu.Add "Label", HeadingTitle(pp.Range.Text)
                    Level2Menu.Add "MenuType", "Heading"
                    Level2Menu.Add "Children", Level2Children
                    Level1Children.Add "Level_2_" & pCount, Level2Menu
                                        
                    If SubHeadingLevel = 2 Then
                        For Each ppp In Level2Range.Paragraphs
                            pCount = pCount + 1
                            If ppp.outlineLevel = 3 Then
                                Set Level3Range = Paperless.SelectHeadingAndContentRange(ppp)
                                Documents(Path).Bookmarks.Add "Level_3_" & pCount, Level3Range
                                Set Level3Menu = New Dictionary
                                Level3Menu.Add "Name", "Level_3_" & pCount
                                Level3Menu.Add "Label", HeadingTitle(ppp.Range.Text)
                                Level3Menu.Add "Path", Path
                                Level3Menu.Add "MenuType", "Heading"
                                Level2Children.Add "Level_3_" & pCount, Level3Menu
                            End If
                        Next ppp
                    End If
                End If
            Next pp
        End If
    Next p
    
    Dim JSON
    JSON = JSONTools.ConvertToJson(Bookmarks)
    Debug.Print JSON
                
    ' Close file and save changes
    Documents(Path).Close SaveChanges:=wdSaveChanges
            
    Set AddBookmarks = Bookmarks
    
    Exit Function
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description

End Function

Public Function LargestHeading() As Integer
    LargestHeading = 3
      
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Style = "Hat"
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute
        
        .ClearFormatting
        .Replacement.ClearFormatting
        
        If .Found Then LargestHeading = 2
    End With
    
    With Selection.Find
        .ClearFormatting
        .Replacement.ClearFormatting
        .Forward = True
        .Wrap = wdFindContinue
        .Format = True
        .Style = "Pocket"
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
        .Execute
        
        .ClearFormatting
        .Replacement.ClearFormatting
        
        If .Found Then LargestHeading = 1
    End With
End Function

Sub RemoveBookmarks()
    Dim b As Bookmark
    For Each b In ActiveDocument.Bookmarks
        b.Delete
    Next b
End Sub
