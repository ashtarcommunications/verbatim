Attribute VB_Name = "OCR"
Option Explicit

Public Sub PasteOCR()
    On Error GoTo Handler

    #If Mac Then
        Filesystem.RequestFolderAccess "/tmp"

        Dim TesseractPath As String
        
        If Filesystem.FileExists("/opt/local/bin/tesseract") Then
            TesseractPath = "/opt/local/bin/tesseract" ' Macports default
        ElseIf Filesystem.FileExists("/usr/local/opt/tesseract") Then
            TesseractPath = "/usr/local/opt/tesseract"
        ElseIf Filesystem.FileExists("/opt/homebrew/bin/tesseract") Then
            TesseractPath = "/opt/homebrew/bin/tesseract"
        ElseIf Filesystem.FileExists("/usr/local/bin/tesseract") Then
            TesseractPath = "/usr/local/bin/tesseract" ' Homebrew default
        ElseIf Filesystem.FileExists("/usr/bin/tesseract") Then
            TesseractPath = "/usr/bin/tesseract"
        Else
            ' Have to use || true to always get a 0 return code
            TesseractPath = AppleScriptTask("Verbatim.scpt", "RunShellScript", "which tesseract || true")
        End If
        
        If TesseractPath = "" Then
            MsgBox "Tesseract is required for OCR functions, and does not appear to be installed. You can install it with MacPorts or Homebrew. For more information, see the Verbatim documentation."
            Exit Sub
        End If
        
        AppleScriptTask "Verbatim.scpt", "RunShellScript", "screencapture -i /tmp/ocrtemp.png"
        If Filesystem.FileExists("/tmp/ocrtemp.png") = False Then
            MsgBox "Something went wrong taking a screenshot. Make sure you have granted Word permissions for Screen Recording in System Preferences -> Security & Privacy."
            Exit Sub
        End If
        
        AppleScriptTask "Verbatim.scpt", "RunShellScript", TesseractPath & " /tmp/ocrtemp.png /tmp/ocrtemp -l ENG"
        If Filesystem.FileExists("/tmp/ocrtemp.txt") = False Then
            MsgBox "Something went wrong with the OCR. Ensure you have Tesseract installed correctly."
            Exit Sub
        End If
        
        Selection.TypeText Filesystem.ReadFile("/tmp/ocrtemp.txt")
        Filesystem.DeleteFile "/tmp/ocrtemp.png"
        Filesystem.DeleteFile "/tmp/ocrtemp.txt"
    #Else
        Dim SnippingToolPath As String
        Dim C2TPath As String
        Dim cmd As String
        Dim TempImagePath As String
        
        Dim OCRPath As String
        OCRPath = GetSetting("Verbatim", "Plugins", "OCRPath", "")
        If OCRPath <> "" Then
            If Filesystem.FileExists(OCRPath) = False Then
                MsgBox "External OCR program not found. Please check the path to the application in your Verbatim settings, or remove it to use the built-in Windows Snipping Tool."
                Exit Sub
            Else
                CreateObject("WSCript.Shell").Run OCRPath, 0, True
                Exit Sub
            End If
        End If
        
        If Filesystem.FileExists(Environ$("ProgramW6432") _
            & Application.PathSeparator _
            & "Verbatim" _
            & Application.PathSeparator _
            & "Plugins" _
            & Application.PathSeparator _
            & "OCR" _
            & Application.PathSeparator _
            & "Capture2Text_CLI.exe" _
        ) = True Then
            C2TPath = Environ$("ProgramW6432") _
                & Application.PathSeparator _
                & "Verbatim" _
                & Application.PathSeparator _
                & "Plugins" _
                & Application.PathSeparator _
                & "OCR" _
                & Application.PathSeparator _
                & "Capture2Text_CLI.exe"
        ElseIf Filesystem.FileExists(Environ$("ProgramW6432") & Application.PathSeparator & "Capture2Text" & Application.PathSeparator & "Capture2Text_CLI.exe") = True Then
            C2TPath = Environ$("ProgramW6432") & Application.PathSeparator & "Capture2Text" & Application.PathSeparator & "Capture2Text_CLI.exe"
        Else
            MsgBox "Capture2Text must be installed to run OCR. Please see https://paperlessdebate.com/ for more details on how to install."
            Exit Sub
        End If
        
        ' Take a screenshot with the snipping tool - have to try alternate paths because FSO can't always find SnippingTool.exe
        On Error Resume Next
        SnippingToolPath = Environ$("SYSTEMROOT") & Application.PathSeparator & "sysnative" & Application.PathSeparator & "SnippingTool.exe"
        CreateObject("WScript.Shell").Run SnippingToolPath & " /clip", 0, True
        If Err.Number <> 0 Then
            Err.Clear
            SnippingToolPath = Environ$("SYSTEMROOT") & Application.PathSeparator & "System32" & Application.PathSeparator & "SnippingTool.exe"
            CreateObject("WScript.Shell").Run SnippingToolPath & " /clip", 0, True
        End If
        If Err.Number <> 0 Then
            Err.Clear
            CreateObject("WScript.Shell").Run "SnippingTool.exe" & " /clip", 0, True
        End If
        If Err.Number <> 0 Then
            MsgBox "Error running the Windows Snipping Tool. Is it installed?"
            Exit Sub
        End If
        On Error GoTo Handler
        
        ' Save screenshot from clipboard to temp file
        TempImagePath = Environ$("TEMP") & Application.PathSeparator & "ocrtemp.jpg"
        cmd = "$img = get-clipboard -format image; $img.save('" & TempImagePath & "');"
        CreateObject("WScript.Shell").Run "powershell -command " & cmd, 0, True
        
        ' OCR temp file
        cmd = """" & C2TPath & """ --clipboard -i """ & TempImagePath & """"
        CreateObject("WScript.Shell").Run cmd, 0, True
        
        ' Delete temp file
        Filesystem.DeleteFile (TempImagePath)
        
        ' Paste OCR from clipboard
        Selection.Paste
    #End If
    
    Exit Sub
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub
