Attribute VB_Name = "Audio"
'@IgnoreModule ProcedureNotUsed, UnassignedVariableUsage, VariableNotAssigned
Option Explicit

Public Sub StartRecord(ByVal BPS As BitsPerSec, ByVal SPS As SamplesPerSec, ByVal Mode As Channels)
    Dim retStr As String
    Dim cBack As Long
    Dim BytesPerSec As Long
    
    On Error GoTo Handler
        
    ' Save instead if already recording
    If Audio.RecordStatus = "recording" Then Audio.SaveRecord
    
    #If Mac Then
        AppleScriptTask "Verbatim.scpt", "StartRecord", ""
    #Else
        ' mciSendString appears to be ignoring the parameters and always recording at 88kbps
        retStr = Space$(128)
        BytesPerSec = (Mode * BPS * SPS) / 8
        mciSendString "open new type waveaudio alias capture", retStr, 128, cBack
        mciSendString "set capture time format milliseconds" & _
          " bitspersample " & CStr(BPS) & _
          " samplespersec " & CStr(SPS) & _
          " channels " & CStr(Mode) & _
          " bytespersec " & CStr(BytesPerSec) & _
          " alignment 4", retStr, 128, cBack
        mciSendString "record capture", retStr, 128, cBack
    #End If
    
    MsgBox "Recording started. Press the button again to stop."
    
    Exit Sub
    
Handler:
    MsgBox "Error " & Err.Number & ": " & Err.Description
    If Audio.RecordStatus = "recording" Then
        Audio.SaveRecord
    End If
    Globals.RecordAudioToggle = False
    Ribbon.RefreshRibbon
    
End Sub

Public Sub SaveRecord()
    Dim retStr As String
    Dim cBack As Long
    
    Dim AudioDir As String
    Dim Filename As String
    Dim FilePath As String
    
    On Error GoTo Handler
    
    ' Get Audio recording directory from settings or use desktop
    AudioDir = GetSetting("Verbatim", "Paperless", "AudioDir", "")
    If Filesystem.FolderExists(AudioDir) = False Then
        #If Mac Then
            FilePath = "/Users/" & Environ("USER") & "/Desktop"
        #Else
            FilePath = CStr(Environ$("USERPROFILE")) & "\Desktop"
        #End If
    Else
        FilePath = AudioDir
    End If

    ' Ensure a trailing separator
    If Right$(FilePath, 1) <> Application.PathSeparator Then FilePath = FilePath & Application.PathSeparator
             
GetFileName:
    Filename = InputBox("Please enter a name for your saved audio file. It will be saved to the following directory:" _
        & vbCrLf & "(Configurable in Settings)" & vbCrLf & FilePath, _
        "Save Audio Recording", _
        "Recording " & Format$(Now, "m d yyyy hmmAMPM"))
    
    ' Exit if no file name or user pressed Cancel, recording is still active
    If Filename = "" Then
        Globals.RecordAudioToggle = True
        Exit Sub
    End If
    
    ' Clean up filename and ensure correct extension
    Filename = Strings.OnlyAlphaNumericChars(Filename)
    #If Mac Then
        If Right(Filename, 4) <> ".m4a" Then Filename = Filename & ".m4a"
    #Else
        If Right$(Filename, 4) <> ".wav" Then Filename = Filename & ".wav"
    #End If
    Filename = FilePath & Filename
    
    ' Test if file exists
    If Filesystem.FileExists(Filename) = True Then
        If MsgBox("File exists. Overwrite?", vbYesNo) = vbNo Then GoTo GetFileName
    End If
    
    #If Mac Then
        AppleScriptTask "Verbatim.scpt", "SaveRecord", Filename
    #Else
        ' Enclose string in quotes for passing to mciSendString
        Filename = """" & Filename & """"
    
        ' Stop and save capture
        retStr = Space$(128)
        mciSendString "stop capture", retStr, 128, cBack
        mciSendString "save capture " & Filename, retStr, 128, cBack
        mciSendString "close capture", retStr, 128, cBack
    #End If
    
    MsgBox "Recording Saved as:" & vbCrLf & Filename, vbOKOnly
    Ribbon.RefreshRibbon
    
    Exit Sub
    
Handler:
    RecordAudioToggle = False
    #If Mac Then
        ' Do Nothing
    #Else
        If Audio.RecordStatus = "recording" Then
            mciSendString "stop capture", retStr, 128, cBack
        End If
    #End If
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Ribbon.RefreshRibbon
End Sub

Public Function RecordStatus() As String
    #If Mac Then
        If Globals.RecordAudioToggle = True Then
            RecordStatus = "recording"
        Else
            RecordStatus = "off"
        End If
    #Else
        Dim retStr As String
        Dim cBack As Long
        
        retStr = Space$(128)
        mciSendString "status capture mode", retStr, 128, cBack
        RecordStatus = retStr
    #End If
End Function

'@Ignore ParameterNotUsed
Public Sub RecordAudio(ByVal c As IRibbonControl, ByVal pressed As Boolean)
    On Error GoTo Handler
    
    If pressed Then
        ' Start recording - Mac ignores the parameters
        Audio.StartRecord Bits8, Samples8000, Mono
        Globals.RecordAudioToggle = True
    Else
        Globals.RecordAudioToggle = False
        
        ' Stop and save recording
        Audio.SaveRecord
    End If
    
    Ribbon.RefreshRibbon
    
    Exit Sub
    
Handler:
    Globals.RecordAudioToggle = False
    MsgBox "Error " & Err.Number & ": " & Err.Description
    Ribbon.RefreshRibbon
End Sub


