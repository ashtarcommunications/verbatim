Attribute VB_Name = "VirtualTub"
Option Explicit

'Globals to ensure menu ID's, button ID's etc. increment correctly
Public VTubXMLDoc As String
Public VTubMenuIDNumber As Long
Public VTubSplitIDNumber As Long
Public VTubButtonIDNumber As Long
Public VTubCurrentFileNumber As Long
Public VTubDepth As Long
Public VTubMaxDepth As Long
Public VTubFileCount As Long
Public VTubLastModified As Date

Public ProgressForm As New frmProgress

Sub GetVTubContent(control As IRibbonControl, ByRef returnedVal)
'Get content for dynamic menu from XML file

    Dim VTubPath As String
    Dim VTubXMLFileName As String
    Dim VTubFolder As Scripting.Folder
    Dim FileNumber As Integer
    
    'Skip Errors
    On Error Resume Next
        
    'Get VTubPath from Settings and make sure it exists
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath", "")
    If VTubPath = "" Or VTubPath = "?" Then
        If MsgBox("You haven't configured a VTub location in the Verbatim settings. Open Settings?", vbYesNo) = vbYes Then
            UI.ShowForm ("Settings")
            Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    'Append trailing \ if missing
    If Right(VTubPath, 1) <> Application.PathSeparator Then VTubPath = VTubPath & Application.PathSeparator
    
    'Check if VTubXMLFile exists
    VTubXMLFileName = VTubPath & "VTub.xml"
    If Filesystem.FileExists(VTubXMLFileName) = False Then
        'If no XML file, return a button to create it
        returnedVal = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"
        returnedVal = returnedVal & "<button id=""CreateVTub"" label=""Create VTub"" onAction=""VirtualTub.VTubCreateButton"" imageMso=""_3DSurfaceMaterialClassic""" & " />"
        returnedVal = returnedVal & "</menu>"
        Exit Sub
    End If

    'If VTubRefreshPrompt is turned on, check if Tub is out of date by comparing date modified of files to XML file
    'Have to loop all files becuase Word changes folder modified date when opening docs
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

    'Open and read the XML file
    FileNumber = FreeFile
    Open VTubXMLFileName For Input As #FileNumber
    returnedVal = Input$(LOF(FileNumber), FileNumber)
    Close #FileNumber
    
    Set VTubFolder = Nothing
    Exit Sub

Handler:
    Set VTubFolder = Nothing
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Sub VTubRefreshButton(control As IRibbonControl)
    VirtualTub.VTubRefresh
End Sub
Sub VTubCreateButton(control As IRibbonControl)
    VirtualTub.VTubCreate
End Sub
Sub VTubSettingsButton(control As IRibbonControl)
    UI.ShowForm ("Settings")
End Sub

Sub VTubCreate()

    Dim VTubPath As String
    Dim VTubFolder As Scripting.Folder
    Dim FSO As Scripting.FileSystemObject
    
    On Error GoTo Handler
    
    ' Get VTubPath from settings
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath")
    
    ' Append trailing \ if missing
    If Right(VTubPath, 1) <> Application.PathSeparator Then VTubPath = VTubPath & Application.PathSeparator
    
    ' Check if XML already exists - prompt to refresh instead or delete file
    If Filesystem.FileExists(VTubPath & "VTub.xml") = True Then
        If MsgBox("VTub already exists - you can update it with the ""Refresh"" button in the VTub menu. Recreate from scratch instead?", vbYesNo + vbDefaultButton2) = vbNo Then
            Set FSO = Nothing
            Exit Sub
        Else
            FSO.DeleteFile (VTubPath & "VTub.xml")
        End If
    End If
    
    ' Set initial folder as the top level of the VTub
    VTubFolder = Filesystem.GetFolder(VTubPath)
    
    ' Reset counters and get initial count of files in the VTub
    VTubFileCount = 0
    VTubDepth = 0
    VTubMaxDepth = 0
    VirtualTub.VTubFileCounterRecursion Folder:=VTubFolder
    
    ' Check if large number of files
    If VTubFileCount > 20 Then
        If MsgBox("You have a large number of files (>20) in the VTub. This could take a few minutes - okay?", vbYesNo, "You sure?") = vbNo Then Exit Sub
    End If
    
    ' Check if more than 2 folder levels
    If VTubMaxDepth > 2 Then MsgBox "VTub can only handle one level of subfolders - files deeper than one subfolder will be ignored.", vbOKOnly
        
    ' Initialize XML Doc with root node
    VTubXMLDoc = "<menu xmlns=""http://schemas.microsoft.com/office/2006/01/customui"">"
       
    ' Reset counters
    VTubMenuIDNumber = 0
    VTubSplitIDNumber = 0
    VTubButtonIDNumber = 0
    VTubCurrentFileNumber = 0
    VTubDepth = 0
    
    ' Show progress bar
    Set ProgressForm = New frmProgress
    ProgressForm.Caption = "Creating VTub..."
    ProgressForm.lblProgress.Width = 0
    ProgressForm.lblCaption.Caption = "File 0 of " & VTubFileCount
    ProgressForm.Show

    ' Seed Recursion with VTubFolder and the XML root node
    VirtualTub.VTubXMLRecursion Folder:=VTubFolder

    ' Trap for cancel button during recursion
    If ProgressForm.Visible = False Then Exit Sub
    
    ' Update progress form as complete
    ProgressForm.lblCaption.Caption = "Processing complete."
    ProgressForm.lblFile.Caption = ""
    ProgressForm.lblProgress.Width = ProgressForm.fProgress.Width - 6

    ' Add standard buttons to the XMLDoc
    VirtualTub.AddDefaultButtons
    
    ' Close top-level menu
    VTubXMLDoc = VTubXMLDoc & "</menu>"

    'Save XML doc
    VTubXMLDoc.Save (VTubPath & "VTub.xml")

    'Clean up
    Set VTubXMLDoc = Nothing
    Set VTubRootNode = Nothing
    Set FSO = Nothing
    
    Unload ProgressForm
    Set ProgressForm = Nothing
    
    'Refresh ribbon and notify
    Ribbon.RefreshRibbon
    MsgBox "VTub successfully created!" & vbCrLf & vbCrLf & "If you get an error when you click OK that ""The document is too large to save. Delete some text before saving."", you can ignore it - it's a bug in Word and won't affect the VTub."
    
    Exit Sub

Handler:
    Set VTubXMLDoc = Nothing
    Set VTubRootNode = Nothing
    Set FSO = Nothing
    MsgBox "Error " & Err.number & ": " & Err.Description
    
End Sub

Private Sub VTubXMLRecursion(Folder As Scripting.Folder, Parent As MSXML2.IXMLDOMElement)

    Dim Menu As String
    Dim Subfolders
    Dim Subfolder As Scripting.Folder
    Dim Files As Scripting.Files
    Dim f As Scripting.File
    Dim ProgressPct As Double
    
    Dim MenuElems
    Dim m
    Dim Refresh As Boolean
    
    'Turn on error-checking
    On Error GoTo Handler
    
    'Increment depth counter
    VTubDepth = VTubDepth + 1
    
    Subfolders = Filesystem.GetSubfoldersInFolder(Folder)
    
    'Iterate through each subfolder in first depth level - XML menu is limited to 5 levels and we need 3 for the file contents
    If VTubDepth < 2 Then
        For Each Subfolder In Subfolders
                     
            'Loop all "menu" elements in the XMLDoc - if creating from scratch, it will be empty.
            'Otherwise, a match means we're refreshing a single node
            Refresh = False
            Set MenuElems = VTubXMLDoc.SelectNodes("//r:menu")
            For Each m In MenuElems
                If m.Attributes.length = 4 Then '4 attributes means menu element is a subfolder or file
                    If m.Attributes(3).Text = Subfolder.Path Then 'Get path from the tag attribute
                        Call VirtualTub.VTubXMLRecursion(Subfolder, m) 'Recurse with the existing node for the subfolder
                        Refresh = True
                    End If
                End If
            Next m
                                 
            'No match found, so we're creating from scratch
            If Refresh = False Then
                'Create a menu node for the folder
                VTubMenuIDNumber = VTubMenuIDNumber + 1 'Increment Menu number to ensure a unique ID
                'Have to use createNode on an IXMLDOMElement object to overload with the NamespaceURI and avoid empty xmlns attributes
                Set Menu = VTubXMLDoc.createNode("element", "menu", "http://schemas.microsoft.com/office/2006/01/customui")
                Parent.appendChild Menu
                Menu.setAttribute "id", "Menu" & VTubMenuIDNumber
                Menu.setAttribute "label", Subfolder.Name
                Menu.setAttribute "imageMso", "Folder"
                Menu.setAttribute "tag", Subfolder.Path
                
                'Reseed the recursion macro with the new subfolder node - this ensures a loop to the bottom level
                Call VirtualTub.VTubXMLRecursion(Subfolder, Menu)
            End If
        Next Subfolder
    End If
    
    'Initialize files collection in current folder
    Set Files = Folder.Files
    
    'Iterate through each file in the folder
    For Each f In Files
        'Check the file is a Word docx and not a temp file
        If Right(f.Name, 4) = "docx" And Left(f.Name, 1) <> "~" Then
            
            'Increment the Progress Bar
            If ProgressForm.Visible = False Then Exit Sub 'Error trap for cancel button
            VTubCurrentFileNumber = VTubCurrentFileNumber + 1
            ProgressPct = VTubCurrentFileNumber / VTubFileCount
            ProgressForm.lblCaption.Caption = Str(Round(ProgressPct * 100, 0)) & "% - " & "Processing File " & VTubCurrentFileNumber & " of " & VTubFileCount
            ProgressForm.lblFile.Caption = "Processing " & f.Name
            ProgressForm.lblProgress.Width = ProgressPct * ProgressForm.fProgress.Width
            If ProgressForm.lblProgress.Width > 15 Then ProgressForm.lblProgress.Width = ProgressForm.lblProgress.Width - 15
            
            DoEvents 'Necessary for Progress form to update
                
            'Loop all menu elements in the XMLDoc and compare paths - if a match, we're refreshing a single node
            Refresh = False
            Set MenuElems = VTubXMLDoc.SelectNodes("//r:menu")
            For Each m In MenuElems
                If m.Attributes.length = 4 Then '4 attributes means menu element is a subfolder or file
                    If Split(m.Attributes(3).Text, "!#!")(0) = f.Path Then 'Recover path by splitting tag attribute
                        Refresh = True
                        If Split(m.Attributes(3).Text, "!#!")(1) < f.DateLastModified Then 'Only update if the file has been modified
                            Set Menu = VTubXMLDoc.createNode("element", "menu", "http://schemas.microsoft.com/office/2006/01/customui")
                            VTubMenuIDNumber = VTubMenuIDNumber + 1 'Increment Menu number to ensure a unique ID
                            Menu.setAttribute "id", "Menu" & VTubMenuIDNumber
                            Menu.setAttribute "label", f.Name
                            Menu.setAttribute "imageMso", "FileSaveAsWordDocx"
                            Set Menu = VTubProcessFile(f.Path, Menu) 'Reprocess file
                            Menu.setAttribute "tag", f.Path & "!#!" & f.DateLastModified 'Save path and date modified - uses !#! as a delimiter
                            
                            'Replace the old node
                            m.ParentNode.replaceChild Menu, m
                            
                        End If
                    End If
                End If
            Next m
                
            'No match found, we're creating file node from scratch
            If Refresh = False Then
                'Create a menu node for the file
                Set Menu = VTubXMLDoc.createNode("element", "menu", "http://schemas.microsoft.com/office/2006/01/customui")
                VTubMenuIDNumber = VTubMenuIDNumber + 1 'Increment Menu number to ensure a unique ID
                Menu.setAttribute "id", "Menu" & VTubMenuIDNumber
                Menu.setAttribute "label", f.Name
                Menu.setAttribute "imageMso", "FileSaveAsWordDocx"
        
                'Process the file and update the menu node, then append it
                Set Menu = VTubProcessFile(f.Path, Menu)
                Menu.setAttribute "tag", f.Path & "!#!" & f.DateLastModified
                Parent.appendChild Menu
            End If
        End If
    Next f
    
    'Clean up
    Set Files = Nothing
    Set Menu = Nothing
    Set MenuElems = Nothing
    Exit Sub
    
Handler:
    Set Files = Nothing
    Set Menu = Nothing
    Set MenuElems = Nothing
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Function VTubProcessFile(FileName As String, FileElement As MSXML2.IXMLDOMElement) As MSXML2.IXMLDOMElement

    Dim SubMenu As MSXML2.IXMLDOMElement
    Dim Button As MSXML2.IXMLDOMElement
    Dim SplitButton As MSXML2.IXMLDOMElement
    
    Dim PocketBMRange As Range
    Dim HatBMRange As Range
    Dim BlockBMRange As Range
    
    Dim PocketBMName As String
    Dim HatBMName As String
    Dim BlockBMName As String
    
    Dim PocketOpen As Boolean
    Dim HatOpen As Boolean
    Dim BlockOpen As Boolean
    
    Dim p As Paragraph
    Dim pCount As Long
    Dim pFix As String
      
    On Error GoTo Handler
      
    'Open the file in the background and activate it
    Documents.Open FileName:=FileName, Visible:=False
    Documents(FileName).Activate
    
    'Delete all bookmarks
    Call VirtualTub.RemoveBookmarks
    
    'Initialize bookmark ranges
    Set PocketBMRange = ActiveDocument.Range
    Set HatBMRange = ActiveDocument.Range
    Set BlockBMRange = ActiveDocument.Range
    
    'Open a new bookmark if first paragraph is a heading
    If Documents(FileName).Paragraphs(1).outlineLevel = wdOutlineLevel1 Then
        PocketOpen = True
    Else
        PocketOpen = False
    End If
    
    If Documents(FileName).Paragraphs(1).outlineLevel = wdOutlineLevel2 Then
        HatOpen = True
    Else
        HatOpen = False
    End If
    
    If Documents(FileName).Paragraphs(1).outlineLevel = wdOutlineLevel3 Then
        BlockOpen = True
    Else
        BlockOpen = False
    End If
    
    'Initialize bookmark names
    pCount = 0
    PocketBMName = "PocketBM1"
    HatBMName = "HatBM1"
    BlockBMName = "BlockBM1"
    
    'Loop all paragraphs
    For Each p In Documents(FileName).Paragraphs
        pCount = pCount + 1
        
        'If end of file, close all bookmarks
        If p.Range.End = Documents(FileName).Range.End Then
            PocketBMRange.End = Documents(FileName).Range.End
            HatBMRange.End = Documents(FileName).Range.End
            BlockBMRange.End = Documents(FileName).Range.End
            If PocketOpen = True Then Documents(FileName).Bookmarks.Add PocketBMName, PocketBMRange
            If HatOpen = True Then Documents(FileName).Bookmarks.Add HatBMName, HatBMRange
            If BlockOpen = True Then Documents(FileName).Bookmarks.Add BlockBMName, BlockBMRange
        End If
        
        'Process depending on outline level
        Select Case p.outlineLevel
        
        Case Is = 1 'Pocket
        
            If Len(p.Range.Text) > 0 Then
    
                'Close open bookmarks
                If PocketOpen = True Then
                    PocketBMRange.End = p.Range.Start
                    Documents(FileName).Bookmarks.Add PocketBMName, PocketBMRange
                    PocketOpen = False
                End If
                If HatOpen = True Then
                    HatBMRange.End = p.Range.Start
                    Documents(FileName).Bookmarks.Add HatBMName, HatBMRange
                    HatOpen = False
                End If
                If BlockOpen = True Then
                    BlockBMRange.End = p.Range.Start
                    Documents(FileName).Bookmarks.Add BlockBMName, BlockBMRange
                    BlockOpen = False
                End If
                
                'Start a new bookmark
                PocketBMRange.Start = p.Range.Start
                PocketBMName = "PocketBM" & pCount
                PocketOpen = True
            
                'Clean text and ensure a non-zero string
                pFix = Trim(OnlySafeChars(Replace(p.Range.Text, Chr(151), "-")))
                If Len(pFix) > 1000 Then pFix = Left(pFix, 1000) 'Limit length to 1000 characters to avoid breaking XML
                If pFix = "" Then pFix = "-"
    
                'Increment Counters
                VTubMenuIDNumber = VTubMenuIDNumber + 1
                VTubSplitIDNumber = VTubSplitIDNumber + 1
                VTubButtonIDNumber = VTubButtonIDNumber + 1
                
                'Make and append buttons - for pockets, it will always be appended directly to file node
                Set SplitButton = VTubXMLDoc.createNode("element", "splitButton", "http://schemas.microsoft.com/office/2006/01/customui")
                FileElement.appendChild SplitButton
                SplitButton.setAttribute "id", "splitButton" & VTubSplitIDNumber
                Set Button = VTubXMLDoc.createNode("element", "button", "http://schemas.microsoft.com/office/2006/01/customui")
                SplitButton.appendChild Button
                Button.setAttribute "id", "Button" & VTubButtonIDNumber
                Button.setAttribute "label", pFix
                Button.setAttribute "tag", FileName & "!#!" & PocketBMName
                Button.setAttribute "onAction", "VirtualTub.VTubInsertBookmark"
                Button.setAttribute "imageMso", "ExportTextFile"
                Set SubMenu = VTubXMLDoc.createNode("element", "menu", "http://schemas.microsoft.com/office/2006/01/customui")
                SplitButton.appendChild SubMenu
                SubMenu.setAttribute "id", "Menu" & VTubMenuIDNumber
            End If
                        
        Case Is = 2 'Hat
            If Len(p.Range.Text) > 0 Then
            
                'Close open bookmarks
                If HatOpen = True Then
                    HatBMRange.End = p.Range.Start
                    Documents(FileName).Bookmarks.Add HatBMName, HatBMRange
                    HatOpen = False
                End If
                If BlockOpen = True Then
                    BlockBMRange.End = p.Range.Start
                    Documents(FileName).Bookmarks.Add BlockBMName, BlockBMRange
                    BlockOpen = False
                End If
                                                    
                'Start a new bookmark
                HatBMRange.Start = p.Range.Start
                HatBMName = "HatBM" & pCount
                HatOpen = True
            
                'Clean text and ensure a non-zero string
                pFix = Trim(OnlySafeChars(Replace(p.Range.Text, Chr(151), "-")))
                If Len(pFix) > 1000 Then pFix = Left(pFix, 1000)
                If pFix = "" Then pFix = "-"
                
                'Increment Counters
                VTubMenuIDNumber = VTubMenuIDNumber + 1
                VTubSplitIDNumber = VTubSplitIDNumber + 1
                VTubButtonIDNumber = VTubButtonIDNumber + 1
                
                'Make and append buttons
                Set SplitButton = VTubXMLDoc.createNode("element", "splitButton", "http://schemas.microsoft.com/office/2006/01/customui")
                
                If FileElement.HasChildNodes = True Then 'If child nodes, then there's already content
                    If PocketOpen = True Then 'If Pocket is open, then the last child will be a pocket, so append to that
                        FileElement.LastChild.LastChild.appendChild SplitButton
                    Else 'If not a pocket, then the Hat is the top level and append to the file element instead
                        FileElement.appendChild SplitButton
                    End If
                Else 'If no child nodes, append to the file element
                    FileElement.appendChild SplitButton
                End If
                
                SplitButton.setAttribute "id", "splitButton" & VTubSplitIDNumber
                Set Button = VTubXMLDoc.createNode("element", "button", "http://schemas.microsoft.com/office/2006/01/customui")
                SplitButton.appendChild Button
                Button.setAttribute "id", "Button" & VTubButtonIDNumber
                Button.setAttribute "label", pFix
                Button.setAttribute "tag", FileName & "!#!" & HatBMName
                Button.setAttribute "onAction", "VirtualTub.VTubInsertBookmark"
                Button.setAttribute "imageMso", "ExportTextFile"
                Set SubMenu = VTubXMLDoc.createNode("element", "menu", "http://schemas.microsoft.com/office/2006/01/customui")
                SplitButton.appendChild SubMenu
                SubMenu.setAttribute "id", "Menu" & VTubMenuIDNumber
            End If
        
        Case Is = 3 'Block
            If Len(p.Range.Text) > 0 Then
        
                'Close open bookmarks
                If BlockOpen = True Then
                    BlockBMRange.End = p.Range.Start
                    Documents(FileName).Bookmarks.Add BlockBMName, BlockBMRange
                    BlockOpen = False
                End If
                
                'Start a new bookmark
                BlockBMRange.Start = p.Range.Start
                BlockBMName = "BlockBM" & pCount
                BlockOpen = True
            
                'Clean text and ensure a non-zero string
                pFix = Trim(OnlySafeChars(Replace(p.Range.Text, Chr(151), "-")))
                If Len(pFix) > 1000 Then pFix = Left(pFix, 1000)
                If pFix = "" Then pFix = "-"
                
                'Create a button, append to deepest level
                Set Button = VTubXMLDoc.createNode("element", "button", "http://schemas.microsoft.com/office/2006/01/customui")
       
                'Looks messy, but necessary to append the block to the correct depth level with splitbuttons
                'Probably a more efficient way to do it with XPath or using siblings
                If FileElement.HasChildNodes = True Then 'If there's child nodes, we have to find deepest level
                    If FileElement.LastChild.HasChildNodes = True Then 'If the last child of the file has children, it's a heading so we have to dig until we run out of children
                        If FileElement.LastChild.LastChild.HasChildNodes = True Then
                            If FileElement.LastChild.LastChild.LastChild.HasChildNodes = True Then
                                If FileElement.LastChild.LastChild.LastChild.LastChild.nodeName = "menu" Then 'If last child is a menu, append to it - if it's not, then append a level higher
                                    FileElement.LastChild.LastChild.LastChild.LastChild.appendChild Button
                                Else
                                    FileElement.LastChild.LastChild.LastChild.appendChild Button
                                End If
                            Else
                                If FileElement.LastChild.LastChild.LastChild.nodeName = "menu" Then
                                    FileElement.LastChild.LastChild.LastChild.appendChild Button
                                Else
                                    FileElement.LastChild.LastChild.appendChild Button
                                End If
                            End If
                        Else
                            If FileElement.LastChild.LastChild.nodeName = "menu" Then
                                FileElement.LastChild.LastChild.appendChild Button
                            Else
                                FileElement.LastChild.appendChild Button
                            End If
                        End If
                    Else
                        If FileElement.LastChild.nodeName = "menu" Then
                            FileElement.LastChild.appendChild Button
                        Else
                            FileElement.appendChild Button
                        End If
                    End If
                Else 'If no child nodes, then append to the file element
                    FileElement.appendChild Button
                End If
                
                VTubButtonIDNumber = VTubButtonIDNumber + 1
                Button.setAttribute "id", "Button" & VTubButtonIDNumber
                Button.setAttribute "label", pFix
                Button.setAttribute "tag", FileName & "!#!" & BlockBMName
                Button.setAttribute "onAction", "VirtualTub.VTubInsertBookmark"
                Button.setAttribute "imageMso", "ExportTextFile"
            End If

        Case Else
            'Do nothing
            
        End Select
        
    Next p
            
    'Close file and save changes
    Documents(FileName).Close SaveChanges:=wdSaveChanges
            
    Set SubMenu = Nothing
    Set Button = Nothing
    Set SplitButton = Nothing
            
    'Return the updated file node
    Set VTubProcessFile = FileElement
    
    Exit Function
    
Handler:
    Set SubMenu = Nothing
    Set Button = Nothing
    Set SplitButton = Nothing
    MsgBox "Error " & Err.number & ": " & Err.Description

End Function

Private Sub VTubFileCounterRecursion(Folder As Scripting.Folder)

    Dim Subfolder As Scripting.Folder
    Dim Files As Scripting.Files
    Dim f As Scripting.File
    
    On Error Resume Next
    
    'Increment the depth level, save the max depth
    VTubDepth = VTubDepth + 1
    If VTubMaxDepth < VTubDepth Then VTubMaxDepth = VTubDepth
    
    'Iterate through each level of subfolders to see how deep it goes
    For Each Subfolder In Folder.Subfolders
        'Reseed the recursion macro with the current subfolder - this ensures a loop to the bottom level
        Call VirtualTub.VTubFileCounterRecursion(Subfolder)
        VTubDepth = VTubDepth - 1 'Decrement depth level when coming out of a subfolder
    Next Subfolder
    
    'Only count files in the first two levels
    If VTubDepth < 3 Then
        
        'Initialize files collection in current folder
        Set Files = Folder.Files
        
        'Iterate through each file in the folder
        For Each f In Files
            'If the file is a Word docx and not a temp file, increment the count
            If Right(f.Name, 4) = "docx" And Left(f.Name, 1) <> "~" Then VTubFileCount = VTubFileCount + 1
            
            'Save the most recent date modified
            If f.DateLastModified > VTubLastModified Then VTubLastModified = f.DateLastModified
        Next f
    
    End If
    
    'Clean up
    Set Files = Nothing

    Exit Sub
    
Handler:
    Set Files = Nothing
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Sub VTubRefresh()

    Dim VTubPath As String
    Dim VTubFolder As Scripting.Folder
    Dim FSO As Scripting.FileSystemObject
    Dim Button As MSXML2.IXMLDOMElement
    
    Dim i
    Dim Elem As MSXML2.IXMLDOMElement
    
    On Error GoTo Handler
    
    'Verify before proceeding
    If MsgBox("Are you sure you want to refresh the VTub?", vbOKCancel) = vbCancel Then Exit Sub
    
    'Get VTubPath from Settings
    VTubPath = GetSetting("Verbatim", "VTub", "VTubPath")
    If Right(VTubPath, 1) <> "\" Then VTubPath = VTubPath & "\" 'Append trailing \ if missing
    
    'Load XML from file
    Set VTubXMLDoc = New DOMDocument60
    VTubXMLDoc.Load (VTubPath & "VTub.xml")
    Set VTubRootNode = VTubXMLDoc.FirstChild
        
    'Delete bottom 4 elements, separator and buttons, to avoid appending below them
    Set Button = VTubRootNode.LastChild
    Button.PreviousSibling.PreviousSibling.PreviousSibling.ParentNode.RemoveChild Button.PreviousSibling.PreviousSibling.PreviousSibling
    Button.PreviousSibling.PreviousSibling.ParentNode.RemoveChild Button.PreviousSibling.PreviousSibling
    Button.PreviousSibling.ParentNode.RemoveChild Button.PreviousSibling
    Button.ParentNode.RemoveChild Button
    
    'Set initial folder as the top level of the VTub
    Set FSO = New Scripting.FileSystemObject
    Set VTubFolder = FSO.GetFolder(VTubPath)

    'Reset counters and get initial count of files in the VTub
    VTubFileCount = 0
    VTubDepth = 0
    Call VirtualTub.VTubFileCounterRecursion(VTubFolder)
    
    'Set the counters to the highest number in the XML Doc to ensure correct incrementing
    VTubXMLDoc.setProperty "SelectionNamespaces", "xmlns:r='http://schemas.microsoft.com/office/2006/01/customui'"
    VTubMenuIDNumber = VTubXMLDoc.SelectNodes("//r:menu").length
    VTubSplitIDNumber = VTubXMLDoc.SelectNodes("//r:splitButton").length
    VTubButtonIDNumber = VTubXMLDoc.SelectNodes("//r:button").length
    VTubCurrentFileNumber = 0
    
    'Show progress bar
    Set ProgressForm = New frmProgress
    ProgressForm.Caption = "Refreshing VTub..."
    ProgressForm.lblProgress.Width = 0
    ProgressForm.lblCaption.Caption = "File 0 of " & VTubFileCount
    ProgressForm.Show
    
    'Seed Recursion with VTubFolder and the XML root node
    Call VirtualTub.VTubXMLRecursion(VTubFolder, VTubRootNode)
    
    'Trap for cancel button during recursion
    If ProgressForm.Visible = False Then Exit Sub
    
    'Update progress form as complete
    ProgressForm.lblCaption.Caption = "Processing complete."
    ProgressForm.lblFile.Caption = ""
    ProgressForm.lblProgress.Width = ProgressForm.fProgress.Width - 6
   
    'Renumber all elements to fix any numbering issues - skip first top-level menu
    For i = 1 To VTubXMLDoc.SelectNodes("//r:menu").length - 1
        Set Elem = VTubXMLDoc.SelectNodes("//r:menu").Item(i)
        Elem.setAttribute "id", "Menu" & i
    Next i
    For i = 0 To VTubXMLDoc.SelectNodes("//r:splitButton").length - 1
        Set Elem = VTubXMLDoc.SelectNodes("//r:splitButton").Item(i)
        Elem.setAttribute "id", "splitButton" & i
    Next i
    For i = 0 To VTubXMLDoc.SelectNodes("//r:button").length - 1
        Set Elem = VTubXMLDoc.SelectNodes("//r:button").Item(i)
        Elem.setAttribute "id", "Button" & i
    Next i
    
    'Add standard buttons back to the bottom
    Call VirtualTub.AddDefaultButtons
    
    'Save the updated XML doc
    VTubXMLDoc.Save (VTubPath & "VTub.xml")
    
    'Clean up
    Set VTubFolder = Nothing
    Set FSO = Nothing
    Set Button = Nothing
    Set VTubXMLDoc = Nothing
    Set VTubRootNode = Nothing
    
    Unload ProgressForm
    Set ProgressForm = Nothing
    
    'Refresh ribbon and notify
    Call Ribbon.RefreshRibbon
    MsgBox "VTub successfully created!" & vbCrLf & vbCrLf & "If you get an error when you click OK that ""The document is too large to save. Delete some text before saving."", you can ignore it - it's a bug in Word and won't affect the VTub."
    
    Exit Sub
    
Handler:
    Set VTubFolder = Nothing
    Set FSO = Nothing
    Set Button = Nothing
    Set VTubXMLDoc = Nothing
    Set VTubRootNode = Nothing
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Private Sub AddDefaultButtons()
    Dim Button As MSXML2.IXMLDOMElement

    Set Button = VTubXMLDoc.createNode("element", "menuSeparator", "http://schemas.microsoft.com/office/2006/01/customui")
    Button.setAttribute "id", "VTubSeparator"
    VTubRootNode.appendChild Button
    
    Set Button = VTubXMLDoc.createNode("element", "button", "http://schemas.microsoft.com/office/2006/01/customui")
    Button.setAttribute "id", "RefreshVTub"
    Button.setAttribute "label", "Refresh VTub"
    Button.setAttribute "onAction", "VirtualTub.VTubRefreshButton"
    Button.setAttribute "imageMso", "AccessRefreshAllLists"
    VTubRootNode.appendChild Button

    Set Button = VTubXMLDoc.createNode("element", "button", "http://schemas.microsoft.com/office/2006/01/customui")
    Button.setAttribute "id", "RecreateVTub"
    Button.setAttribute "label", "Recreate VTub"
    Button.setAttribute "onAction", "VirtualTub.VTubCreateButton"
    Button.setAttribute "imageMso", "_3DSurfaceMaterialClassic"
    VTubRootNode.appendChild Button
    
    Set Button = VTubXMLDoc.createNode("element", "button", "http://schemas.microsoft.com/office/2006/01/customui")
    Button.setAttribute "id", "VTubSettings"
    Button.setAttribute "label", "VTub Settings"
    Button.setAttribute "onAction", "VirtualTub.VTubSettingsButton"
    Button.setAttribute "imageMso", "_3DLightingFlatClassic"
    VTubRootNode.appendChild Button

    Set Button = Nothing

End Sub

Sub RemoveBookmarks()

    Dim b As Bookmark
    For Each b In ActiveDocument.Bookmarks
        b.Delete
    Next b

End Sub

Sub VTubInsertBookmark(control As IRibbonControl)
        
    'Insert bookmark - get the file path and bookmark name by splitting the tag attribute on the !#! delimiter
    Selection.InsertFile Split(control.Tag, "!#!", 2)(0), Split(control.Tag, "!#!", 2)(1)

End Sub

Private Sub TestVTub()
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
    
    Set Folder = GetFolder(GetSetting("Verbatim", "VTub", "VTubPath", ""))
      
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

    RootMenu.Add "FileCount", FileCount
    RootMenu.Add "FolderCount", Folder.Subfolders.Count
    RootMenu.Add "DateLastModified", Folder.DateLastModified

    For i = 1 To Folder.Subfolders.Count
        Set Subfolder = GetFolder(Folder.Subfolders(i))
        Set SubfolderMenu = New Dictionary
        SubfolderMenu.Add "MenuType", "Folder"
        SubfolderMenu.Add "Name", Subfolder.Name
        SubfolderMenu.Add "Path", Subfolder.Path
        
        
        Set Children = New Dictionary
        
        For j = 1 To Subfolder.Files.Count
            If Right(Subfolder.Files(j), 4) = "docx" Then
                Set File = GetFile(Subfolder.Files(j))
                Set FileMenu = New Dictionary
                FileMenu.Add "MenuType", "File"
                FileMenu.Add "Name", File.Name
                FileMenu.Add "Path", File.Path
                
                Set Headings = VirtualTub.AddBookmarks(File.Path)
                
                FileMenu.Add "Children", Headings
                FileMenu.Add "DateLastModified", File.DateLastModified
                
                Children.Add Replace(File.Path, "\", "\\"), FileMenu
            End If
        Next
        
        SubfolderMenu.Add "Children", Children
        SubfolderMenu.Add "DateLastModified", Subfolder.DateLastModified

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
            FileMenu.Add "DateLastModified", File.DateLastModified
            
            RootMenu.Add Replace(File.Path, "\", "\\"), FileMenu
        End If
    Next
        
    JSON = JSONTools.ConvertToJson(RootMenu)
    Debug.Print JSON
       
    'Save file
    Dim VTubFilePath
    Dim OutputFile
    VTubFilePath = "C:\Users\hardy\Desktop\Tub\VTub.json"
    OutputFile = FreeFile
    Open VTubFilePath For Output As #OutputFile
    Print #OutputFile, JSON
    Close #OutputFile
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description

End Sub

Sub RefreshVTub()
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

    Set RootMenu = JSONTools.ParseJson(JSON)
    
    Dim Folder As clsFolder
    Dim FileCount As Long
    FileCount = 0
       
    Set Folder = GetFolder(GetSetting("Verbatim", "VTub", "VTubPath", ""))
      
    For i = 1 To Folder.Subfolders.Count
        Set Subfolder = GetFolder(Folder.Subfolders(i))
        FileCount = FileCount + Subfolder.Files.Count
    Next
    
    FileCount = FileCount + Folder.Files.Count
    
    If (FileCount <> RootMenu("FileCount") Or Folder.Subfolders.Count <> RootMenu("FolderCount")) Then
        If MsgBox("The number of files or folders in your VTub appear to have changed and needs to be rebuilt from scratch. Rebuild now?", vbOKCancel) = vbCancel Then Exit Sub
        VirtualTub.TestVTub
        Exit Sub
    End If
    
    Dim key As Variant
    For Each key In RootMenu.Keys
        Set Menu = RootMenu(key)
        If (Menu("MenuType") = "Folder" And Menu.Exists("Children")) Then
            Set Children = Menu("Children")

            For Each key In Children.Keys
                Set Child = Children(key)
                If Child("MenuType") = "File" Then
                    Set File = GetFile(Child("Path"))
                    If Child("DateLastModified") <> File.DateLastModified Then
                        Child("Children") = AddHeadings(File.Path)
                        Set File = GetFile(File.Path)
                        Child("DateLastModified") = File.DateLastModified
                    End If
                End If
            Next key
            
            Set Folder = GetFolder(Menu("Path"))
            Menu("DateLastModified") = Folder.DateLastModified
        ElseIf Menu("MenuType") = "File" Then
            Set File = GetFile(Menu("Path"))
            If Menu("DateLastModified") <> File.DateLastModified Then
                Menu("Children") = AddHeadings(File.Path)
                Set File = GetFile(File.Path)
                Menu("DateLastModified") = File.DateLastModified
            End If
        End If
    Next key
  
    'Save file
    Dim VTubFilePath
    Dim OutputFile
    VTubFilePath = "C:\Users\hardy\Desktop\Tub\VTub.json"
    OutputFile = FreeFile
    Open VTubFilePath For Output As #OutputFile
    Print #OutputFile, JSON
    Close #OutputFile
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.number & ": " & Err.Description
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
        Set Menu = RootMenu(key)
        xml = xml & ConvertDictionaryToXML(Menu)
        
    Next key
  
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
    MsgBox "Error " & Err.number & ": " & Err.Description
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
    MsgBox "Error " & Err.number & ": " & Err.Description

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
