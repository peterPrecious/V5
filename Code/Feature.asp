<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->


<html>

  <head>
    <meta charset="UTF-8">
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
    <script src="/V5/Inc/jQuery.js"></script>
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
    <script src="/V5/Inc/Functions.js"></script>
    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  </head>


  <body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

    <% Server.Execute vShellHi %>
    <div align="center">
      <center>
    <table border="0" cellpadding="0" cellspacing="0" width="75%">
      <tr>
        <td width="100%">
        <%
          Dim vUrl 
          '...note: vId ends in "." so the translation engine converts it - thus don't need the dot in ".asp"
          vUrl = "..\Features\" & Request.QueryString("vId") & "htm"
          On Error Resume Next
          Server.Execute vUrl
          On Error Goto 0
          If svSecure Then 
            vPage = svCustCluster & ".asp"
          Else  
            vPage = "Welcome.asp"
          End If         
        %> 
        </td>
      </tr>
    </table>
    </center>
    </div>

    <p align="center"><input onclick="history.back()" type="button" value="<%=fPhraH(000257)%>" name="bReturn" id="bReturn" class="button"></p>

    <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>

</html>



