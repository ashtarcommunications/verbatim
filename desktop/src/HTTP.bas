Attribute VB_Name = "HTTP"
Option Explicit

Public Function GetReq(ByRef URL As String) As Dictionary
    On Error GoTo Handler
    
    Dim Response As Dictionary
    Set Response = New Dictionary
        
    #If Mac Then
        Dim Script As String
        Dim Raw As String
        Dim StatusCode As String
        Dim Body As String
        
        ' Get the response from curl as <status_code>\n<body> by munging stdout with awk
        ' -w is the write-out string format, -s is silent
        Script = "curl -w '\n%{http_code}' " _
            & "--cookie 'caselist_token=" & GetSetting("Verbatim", "Caselist", "CaselistToken", "") & "' " _
            & "-H 'Accept: application/json' " _
            & "-s '" & URL & "' | awk 'NR==1{code=$0} NR>1{print} END{if (code) print code}'"
        Raw = AppleScriptTask("Verbatim.scpt", "RunShellScript", Script)
        
        ' Chr(13) = matches \n newline output from curl - should only be one newline in output
        StatusCode = Split(Raw, Chr(13))(0)
        Body = Split(Raw, Chr(13))(1)

        Response.Add "status", StatusCode
        Response.Add "body", JSONTools.ParseJson(Body)
        
        Set GetReq = Response
        
        Set Response = Nothing
    #Else
        Dim HttpReq As Object
        Set HttpReq = CreateObject("MSXML2.ServerXMLHTTP")
        HttpReq.setTimeouts 2000, 10000, 30000, 30000
        HttpReq.Open "GET", URL, False
        HttpReq.setRequestHeader "Accept", "application/json"
        HttpReq.setRequestHeader "Cookie", "caselist_token=" & GetSetting("Verbatim", "Caselist", "CaselistToken", "")
      
        HttpReq.send
        
        Response.Add "status", HttpReq.status
        Response.Add "body", JSONTools.ParseJson(HttpReq.responseText)
             
        Set GetReq = Response

        Set HttpReq = Nothing
        Set Response = Nothing
    #End If

    Exit Function
    
Handler:
    ' Return an empty response if there was a network error
    If Not Response.Exists("status") Then
        Response.Add "status", "Error " & Err.Number & ": " & Err.Description
    End If
    If Not Response.Exists("body") Then
        Response.Add "body", JSONTools.ParseJson("[]")
    End If
    Set GetReq = Response
    
    #If Mac Then
        Set Response = Nothing
    #Else
        Set HttpReq = Nothing
        Set Response = Nothing
    #End If
End Function

Public Function PostReq(ByRef URL As String, ByVal Body As Dictionary) As Dictionary
    On Error GoTo Handler
    
    Dim Response As Dictionary
    Set Response = New Dictionary
        
    Dim JSON As String
    JSON = JSONTools.ConvertToJson(Body)
    
    #If Mac Then
        Dim Script As String
        Dim Cookie As String
        Dim Raw As String
        Dim StatusCode As String
        Dim ResponseBody As String
        
        Cookie = GetSetting("Verbatim", "Caselist", "CaselistToken", "")
        
        ' Uses same output redirection as GET to retrieve status code and response body
        Script = "curl -X POST "
        Script = Script & "-H 'Content-Type: application/json' "
        Script = Script & "-H 'Cookie: caselist_token=" & Cookie & "' "
        Script = Script & "-d '" & JSON & "' "
        Script = Script & "-w '\n%{http_code}' -s "
        Script = Script & "'" & URL & "'"
        Script = Script & " | awk 'NR==1{code=$0} NR>1{print} END{if (code) print code}'"
        
        Raw = AppleScriptTask("Verbatim.scpt", "RunShellScript", Script)
        StatusCode = Split(Raw, Chr(13))(0)
        ResponseBody = Split(Raw, Chr(13))(1)

        Response.Add "status", StatusCode
        Response.Add "body", JSONTools.ParseJson(ResponseBody)
        
        Set PostReq = Response
        
        Set Response = Nothing
    #Else
        Dim HttpReq As Object
        Set HttpReq = CreateObject("MSXML2.ServerXMLHTTP")
        HttpReq.setTimeouts 2000, 10000, 30000, 30000
        HttpReq.Open "POST", URL, False
        HttpReq.setRequestHeader "Accept", "application/json"
        HttpReq.setRequestHeader "Content-Type", "application/json"
        HttpReq.setRequestHeader "Cookie", "caselist_token=" & GetSetting("Verbatim", "Caselist", "CaselistToken", "")
      
        HttpReq.send JSON
        
        Response.Add "status", HttpReq.status
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
    ' Return an empty response if there was a network error
    If Not Response.Exists("status") Then
        Response.Add "status", "Error " & Err.Number & ": " & Err.Description
    End If
    If Not Response.Exists("body") Then
        Response.Add "body", JSONTools.ParseJson("[]")
    End If
    
    Set PostReq = Response
    
    #If Mac Then
        Set Response = Nothing
    #Else
        Set HttpReq = Nothing
        Set Response = Nothing
        Set ResponseBody = Nothing
    #End If
End Function
