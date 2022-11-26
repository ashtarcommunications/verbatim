Attribute VB_Name = "HTTP"
Option Explicit

Public Function GetReq(URL As String) As Dictionary
    On Error GoTo Handler
    
    #If Mac Then
        Dim Script
        Script = "curl '" & URL & "'"
        
        Set GetReq = JSONTools.ParseJson(AppleScriptTask("Verbatim.scpt", "DoShellScript", Script))
    #Else
        Dim HttpReq As MSXML2.ServerXMLHTTP60
        Set HttpReq = New ServerXMLHTTP60
        HttpReq.Open "GET", URL, False
        HttpReq.setRequestHeader "Accept", "application/json"
        HttpReq.setRequestHeader "Cookie", "caselist_token=" & GetSetting("Verbatim", "Caselist", "CaselistToken", "")
      
        HttpReq.send
        
        Dim Response As Dictionary
        Set Response = New Dictionary
        Response.Add "status", HttpReq.Status
        Response.Add "body", JSONTools.ParseJson(HttpReq.responseText)
             
        Set GetReq = Response

        Set HttpReq = Nothing
    #End If

    Exit Function
    
Handler:
    #If Not Mac Then
        Set HttpReq = Nothing
    #End If
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Function

Public Function PostReq(URL As String, Body As Dictionary) As Dictionary
    On Error GoTo Handler
    
    Dim JSON
    JSON = JSONTools.ConvertToJson(Body)
    
    #If Mac Then
        Dim Script
        Script = "do shell script ""curl '" & URL & "'"""
        PostReq = MacScript(Script)
        
        Dim Script
        Script = "curl -X POST "
        Script = Script & "-H 'Content-Type: application/json' "
        Script = Script & "-H 'Cookie: caselist_token=" & GetSetting("Verbatim", "Caselist", "CaselistToken", "") & "' "
        Script = Script & "-d '" & JSON & "' "
        Script = Script & "'" & URL & "'"
        
        Set PostReq = JSONTools.ParseJson(AppleScriptTask("Verbatim.scpt", "DoShellScript", Script))
    #Else
        Dim HttpReq As MSXML2.ServerXMLHTTP60
        Set HttpReq = New ServerXMLHTTP60
        HttpReq.Open "POST", URL, False
        HttpReq.setRequestHeader "Accept", "application/json"
        HttpReq.setRequestHeader "Content-Type", "application/json"
        HttpReq.setRequestHeader "Cookie", "caselist_token=" & GetSetting("Verbatim", "Caselist", "CaselistToken", "")
      
        HttpReq.send JSON
        
        Dim Response As Dictionary
        Set Response = New Dictionary
        Response.Add "status", HttpReq.Status
        Dim ResponseBody As Object
        Set ResponseBody = JSONTools.ParseJson(HttpReq.responseText)
        
        Response.Add "body", ResponseBody
      
        Set PostReq = Response
      
        Set HttpReq = Nothing
        Set Response = Nothing
        Set ResponseBody = Nothing
    #End If

    Exit Function
    
Handler:
    Set HttpReq = Nothing
    Set Response = Nothing
    Set ResponseBody = Nothing
    MsgBox "Error " & Err.Number & ": " & Err.Description
End Function

Public Sub DownloadFile(URL As String, FilePath As Variant)
    #If Mac Then
        MacScript ("do shell script ""curl -o '" & URL & "'" & " " & FilePath & """")
    #Else
        Dim HttpReq As MSXML2.ServerXMLHTTP60
        Set HttpReq = New ServerXMLHTTP60
        HttpReq.Open "GET", URL, False
        HttpReq.send
        Dim FileStream
        Set FileStream = CreateObject("ADODB.Stream")
        FileStream.Open
        FileStream.Type = 1
        FileStream.Write HttpReq.ResponseBody
        FileStream.SaveToFile FilePath, 2 '1 = no overwrite, 2 = overwrite
        FileStream.Close
        Set FileStream = Nothing
    #End If
End Sub
