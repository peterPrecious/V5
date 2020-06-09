<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<html>

<head>
  <title>Error</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <div style="text-align: center; margin: 30px; font-weight: bold;" class="c5">
    <%=Request("vErr")%><br /><br />
    <% 
      '...add if return address button, previous address button or none if vReturn=n
      Dim vReturn
      vReturn = Lcase(Request.QueryString("vReturn"))               
      If Len(vReturn) = 0 Then
        Response.Write "<p align='center'><input onclick='history.back(1)' type='button' value='" & bReturn & "' name='bReturn' class='button'></p>"
      ElseIf vReturn = "close" Then 
        Response.Write "<p align='center'><input onclick=""window.open('', '_parent', '');window.close();"" type='button' value='" & bClose & "' name='bClose' class='button'></p>"
      ElseIf vReturn <> "n" Then 
        Response.Write "<p align='center'><input onclick=""location.href='" & vReturn & "'"" type='button' value='" & bReturn & "' name='bReturn' class='button'></p>"
      End If 
    %>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>
</html>