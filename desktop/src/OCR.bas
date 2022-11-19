Attribute VB_Name = "OCR"
Option Explicit

Public Sub GetOCR()
    On Error GoTo Handler
    #If Mac Then
        ' TODO - Mac version
        MsgBox "OCR not supported on Mac"
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
                MsgBox "External OCR program not found. Please check your settings."
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
        C2TPath = "C:\Users\Aaron\Desktop\Capture2Text\Capture2Text_CLI.exe"
        If C2TPath = vbNullString Or Filesystem.FileExists(C2TPath) = False Then
            MsgBox "Capture2Text must be installed to run OCR. Please see https://paperlessdebate.com/plugins for more details."
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
    #If Not Mac Then
        Set FSO = Nothing
    #End If
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Sub

