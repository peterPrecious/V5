<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <link href="/V5/Inc/<%=Left(svCustId, 4)%>.css" type="text/css" rel="stylesheet">
  <link href="/V5/Inc/Buttons.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <%  Server.Execute vShellHi %> <br>
  <div align="center">
    <table border="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#ECECF9" width="90%">
      <tr>
        <td align="center">
          <h1><%=Request("vMsg")%></h1>
          <a id="aNext" class="butShell" href="javascript:void(0)"><span class="butIcon butPrevious  "></span><input onclick="history.go(-1)" class="butInput" type="button" name="sNext" id="txt04" value="Previous" /> </a>        
          <br>
        </td>
      </tr>
    </table>
    <!--#include virtual = "V5/Inc/Shell_LoLite.asp"-->

</body>

</html>
