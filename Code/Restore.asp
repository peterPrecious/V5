<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% 
  '...Delete Details (used to clear out items when building the plan before updating)
  If Request("Password") = "lesbian" Then
    sOpenCmd
    With oCmd
      .CommandText = "spRestore"
      .Parameters.Append .CreateParameter("@Server",    		adVarChar,  adParamInput,    50, svServer)
    End With
    oCmd.Execute()
    Set oCmd = Nothing
    sCloseDb
    Response.Redirect "Menu.asp"
  End IF
%>

<html>

<head>
  <meta http-equiv="Content-Language" content="en-us">
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <link href="/V5/Inc/Button.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <form method="POST" name="fRestore" action="Restore.asp">
    <table border="0" width="100%" cellspacing="0" cellpadding="2">
      <tr>
        <th>
          <h1><br>Restore Last Nights Backup</h1>
          <p align="left">This utility should only be use by BIG ADMIN to restore last night's data base backup to the <b>Staging Server</b>.&nbsp; Note: the restore typically takes about 10 minutes and the progress of the restore does NOT show on this page. So, once you click <b>Next</b> below, assume the restore has started and check the status of the system in 10 minutes.</p>
          Password: <input type="text" name="Password" size="20"><%=fButton("Next", bNext)%><br><br>(Restoring: <%=svServer%>)<br><br>
        </th>
      </tr>
    </table>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


