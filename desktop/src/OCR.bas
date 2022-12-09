Attribute VB_Name = "OCR"
Option Explicit

Public Sub PasteOCR()
    On Error GoTo Handler
    #If Mac Then
        ' TODO - Mac version
        MsgBox "OCR not supported on Mac"
        
        ' Check tesseract is installed first, then applescript:
        
        'set outPath to "/tmp"
        'set tesseractCmd to (do shell script "zsh -l -c 'which tesseract'")
        'do shell script "screencapture -i " & outPath & "/untitled.png"
        'do shell script tesseractCmd & " " & outPath & "/untitled.png " & outPath & "/output -l jpn"
        'set the_text to (do shell script "cat " & outPath & "/output.txt")
        'set the clipboard to the_text
        'do shell script "rm " & outPath & "/untitled.png " & outPath & "/output.txt"

        
    #Else
        Dim SnippingToolPath As String
    
        Dim FSO As FileSystemObject
        Set FSO = New FileSystemObject
        
        Dim cmd As String
        
        Dim TempImagePath As String
        
        Dim C2TPath As String
        
        Dim ExternalOCR As String
        ExternalOCR = GetSetting("Verbatim", "Plugins", "ExternalOCR", vbNullString)
        If ExternalOCR <> "" Then
            If Filesystem.FileExists(ExternalOCR) = False Then
                MsgBox "External OCR program not found. Please check the path to the application in your Verbatim settings, or remove it to use the built-in Windows Snipping Tool."
                Exit Sub
            Else
                CreateObject("WSCript.Shell").Run ExternalOCR, 0, True
                Exit Sub
            End If
        End If
        
        SnippingToolPath = Environ("SYSTEMROOT") & Application.PathSeparator & "sysnative" & Application.PathSeparator & "SnippingTool.exe"
        
        If Filesystem.FileExists(SnippingToolPath) = False Then
            MsgBox "The Windows Snipping Tool must be installed to run OCR"
            Exit Sub
        End If
        
        C2TPath = GetSetting("Verbatim", "Plugins", "Capture2Text", vbNullString)
        If C2TPath = vbNullString Then
            C2TPath = Environ("ProgramW6432") & Application.PathSeparator & "Capture2Text" & Application.PathSeparator & "Capture2Text_CLI.exe"
        End If
        If Filesystem.FileExists(C2TPath) = False Then
            MsgBox "Capture2Text must be installed to run OCR. Please see https://paperlessdebate.com/ for more details or check the path to the application in your Verbatim settings."
        End If
        
        ' Take a screenshot with the snipping tool
        CreateObject("WSCript.Shell").Run SnippingToolPath & " /clip", 0, True
        
        ' Save screenshot from clipboard to temp file
        TempImagePath = Environ("TEMP") & Application.PathSeparator & "ocrtemp.jpg"
        cmd = "$img = get-clipboard -format image; $img.save('" & TempImagePath & "');"
        CreateObject("WScript.Shell").Run "powershell -command " & cmd, 0, True
        
        ' OCR temp file
        cmd = C2TPath & " --clipboard -i '" & TempImagePath & "'"
        CreateObject("WScript.Shell").Run "powershell -command " & cmd, 0, True
        
        ' Delete temp file
        Filesystem.DeleteFile (TempImagePath)
        
        ' Paste OCR from clipboard
        Selection.Paste
        
        ' Clean up
        Set FSO = Nothing
    #End If
    
    Exit Sub
Handler:
    #If Mac Then
        ' Do Nothing
    #Else
        Set FSO = Nothing
    #End If
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

