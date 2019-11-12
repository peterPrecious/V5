<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/QueryString.asp"-->

<% 
  '...get the parms  
  Dim vUrl
  sGetQueryString
  vId = ""
  sPutQueryString
' sDebugQueryString

	'...values are set in Start.asp
  If Len(vSource) > 0 Then
    vUrl = vSource
	  vUrl = Replace (vUrl, "~1", "&")
	  vUrl = Replace (vUrl, "~2", "=")
	  vUrl = Replace (vUrl, "~3", "?")
  Else  
    vUrl = "/V5/Login.asp?vCust=" & vCust & "&vLang=" & svLang
  End If
%>

<html>

<head>
  <title>SignInErr</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <div style="text-align: center;">
    <h5>
      <% If Not fNoValue(svCustBanner) Then %>
      <img border="0" src="/V5/Images/Logos/<%=svCustBanner%>"><br>
      <% End If %>
      <br><br>
      <!--webbot bot='PurpleText' PREVIEW='You were unable to sign in because'--><%=fPhra(000051)%><br>
      <%=Request("vError")%><br><br></h5>

    <input onclick="location.href = '<%=vUrl%>'" type="button" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button">
    <br><br>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


