<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% 
  '...values can be received by querystring or form but are forward only by querystring
  Dim vNext, vFilter
  
  '...If info received via a querystring (or form get)
  If Len(Request.QueryString("vNext")) > 0 Then
    vNext = Request.QueryString("vNext") & "?" & Request.ServerVariables("QUERY_STRING")

  '...Else if received via a form use filter to strip unnecessary fields
  ElseIf Len(Request.Form("vNext")) > 0 Then
    vNext = Request.Form("vNext") & "?"
    '...see if there's a filter to control what fields to pass through (important as querystrings are limited in length
    vFilter = Request.Form("vFilter")
    For Each vFld in Request.Form
      '...filter fields 
      If Len(vFilter) > 0 Then
        If Instr(vFilter, vFld) > 0 Then
          vNext = vNext & vFld & "=" & Server.UrlEncode(Request(vFld)) & "&"
        End If
      Else
        vNext = vNext & vFld & "=" & Server.UrlEncode(Request(vFld)) & "&"
      End If
    Next
    If Right(vNext, 1) = "&" Then vNext = Left(vNext, Len(vNext)-1)    
  Else
    Response.Redirect "Error.asp?vErr=" & Server.UrlEncode("This service could not continue as it was not properly configured!")
  End If
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Patience</title>
</head>

<body onload="location.href='<%=vNext%>';">

  <% Server.Execute vShellHi %>

  <div style="text-align: center">
    <h1><!--webbot bot='PurpleText' PREVIEW='Please be patient.'--><%=fPhra(000215)%></h1>
    <h2><!--webbot bot='PurpleText' PREVIEW='It can take several minutes for the next page to appear.'--><%=fPhra(000627)%></h2>
    <% If Len(Request("vMsg")) > 0 Then %><p><%=Request("vMsg")%></p><% End If %>
    <p><img border="0" src="../Images/Common/ProgressBar.gif"></p>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

