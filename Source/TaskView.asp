<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_TskH.asp"-->
<!--#include virtual = "V5/Inc/Db_TskD.asp"-->
<!--#include virtual = "V5/Inc/Db_Keys.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Actn.asp"-->
<!--#include virtual = "V5/Inc/Db_Dial.asp"-->
<!--#include virtual = "V5/Inc/Db_Docs.asp"-->
<!--#include virtual = "V5/Inc/Db_Caln.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->

<!--#include virtual = "V5/Inc/Password_Routines.asp"-->
<!--#include virtual = "V5/Inc/ProgramStatusRoutines.asp"-->
<!--#include file = "ModuleStatusRoutines.asp"-->
<!--#include file = "MyWorldCode.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>My Learning</title>
</head>

<body>

  <% 
    Server.Execute vShellHi 
    Dim vNoTasks, vGrid, vBorder, vInfo
    vBorder=0
  %>


  <div align="center">
    <center>
    <table border="0" cellpadding="0" cellspacing="0" width="625">
      <tr>
        <td width="100%" align="center">
        <h5>This is quick view of this task set which is NOT functional.</h5>
        </td>
      </tr>
    </table>
    <table border="1" width="625" style="border-collapse: collapse">
      <% 
        '... before we even start, get member access info and the level 0 values
        sGetMemb (svMembNo)
        Session("MyWorldTree") = ""
        Response.Write fMyWorld (Request.QueryString("vTskH_AcctId"), Request.QueryString("vTskH_Id"), "template")
        Session("MyWorldTree") = "" 
      %>
    </table>
    </center>
    <form method="POST" action="taskview.asp" webbot-action="--WEBBOT-SELF--">
      <!--webbot bot="SaveResults" U-File="../_private/form_results.csv" S-Format="TEXT/CSV" S-Label-Fields="TRUE" startspan --><input TYPE="hidden" NAME="VTI-GROUP" VALUE="0"><!--webbot bot="SaveResults" endspan i-checksum="43374" -->
      <p><input onclick="location.href='TaskEdit1.asp'" type="button" value="Return" name="B3" class="button"></p>
    </form>
    <p>&nbsp;</div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>