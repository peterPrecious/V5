<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Utf8.asp"-->

<%
  '...IIS comes here when this URL is not in the root of the web (like /V5), ie could be //vubiz.com/CFIB
  '   Check if this folder is in ChAccess

'  vUrl = "404;//déléguéssyndicaux"
' "//localhost/V5/Redirect.asp?404;//déléguéssyndicaux"

  Dim oFs, vUrl, vFrom, vFile
  bDebug = True
' bDebug = False

  '...parse the URL to get possible ChAccess folder/page
  vUrl = Request.QueryString.Item 
  If bDebug Then Response.Write "<br>Input: " & vUrl & "<br>"    

  '...IIS can add port 80 to the redirection
  vUrl = Replace(vUrl, ":80", "")
  If bDebug Then Response.Write "<br>" & vUrl & "<br>"    

  If Left(vUrl, 4) = "404;" Then
    i = InStr(vUrl, svServer)
    If bDebug Then Response.Write "<br>" & i & "<br>"    
    vFrom = Mid(vUrl, 5) '...save for error page


    vUrl = "/ChAccess" + Mid(vUrl, i + Len(svServer)) 

    On Error Resume Next
    vFile = Server.MapPath(vUrl)      
    On Error Goto 0
    If Err.Number = 0 Then
      If bDebug Then Response.Write "<br>" & vUrl & "<br>"    
      If i > 0 Then  
        Set oFs = CreateObject("Scripting.FileSystemObject")
        '...looking for a folder or file (containing a ".")
        If Instr(vUrl, ".") = 0 Then
          If oFs.FolderExists(vFile) Then Response.Redirect vUrl '...folder
        Else
          If oFs.FileExists(vFile)   Then Response.Redirect vUrl '...file
        End If
      End If
    End If
  End If   
   
  '...if no folder or file then assume it's an invalid URL
'  Response.Redirect "ErrorPage.asp?vFrom=" & vFrom



%>

<html>
<head>
  <title>:: URL Redirect</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <!--    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">-->
  <!--    <meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />-->
</head>
<body></body>
</html>
