<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
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
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->

<!--#include virtual = "V5/Inc/Certificate.asp"-->
<!--#include virtual = "V5/Inc/Password_Routines.asp"-->
<!--#include virtual = "V5/Inc/ProgramStatusRoutines.asp"-->
<!--#include file = "ModuleStatusRoutines.asp"-->

<!--#include file = "MyWorld2Code.asp"-->
<!--#include file = "MyWorld2CodeRoutines.asp"-->

<html>

<head>
  <title>My World 2</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script src="/V5/Inc/Launch.js"></script>
  <script src="/V5/Inc/LaunchObjects.js"></script>
  <script>
    function fAlert() {
      var vPhrase = "/*--{[--*/You have no more attempts available for this assessment./*--]}--*/"
      alert(vPhrase);
    }

    function getProgramData(vProgId, vMembId) {
      var calledPrograms = new Array();      
      var calledOk = false;
      for (i=0; i < calledPrograms.length; i++) {
        if (calledPrograms[i].value = vProgId) { 
        	calledOk = true;
        }
      }   
      if (!calledOk) {
        calledPrograms[calledPrograms.length++] = vProgId;
        var vParams = 'vFunction=modules&vProgId=' + vProgId + '&vMembId=' + vMembId;
        var vWs = WebService("MyWorld2Code_ws.asp", vParams)
        if (vWs == "error") {
          vWs = "<font color='red'>Web Service could not process: '" + vParams + "'.<br><a onclick='location.reload();' href='#'><font color='red'>Click here to refresh this page then click the above Program link again.</font></a><br> Contact VUBIZ Systems if no content appears.</font>";
        }
        else if (vWs == "err") {
          vWs = "<font color='red'>Web Service could not find Program: '" + vParams + "'.<br><a onclick='location.reload();' href='#'><font color='red'>Click here to refresh this page then click the above Program link again.</font></a><br> Contact VUBIZ Systems if no content appears.</font>";
        }

  	    document.getElementById("div_" + vProgId).innerHTML = vWs; 
    	} 
   	  toggle('div_' + vProgId);
    }

  </script>

  <style>
    td {
      padding: 5px;
      border: 0;
      vertical-align: top;
      text-align: left !important;
    }
  </style>
</head>

<body>

  <% 
    Server.Execute vShellHi 

  response.write "2<br>" ' *****************************************************************************************************

    Dim vNoTasks, vGrid, vBorder, vInfo
 
    '...If there's a single TskH_Id then display My Learning 
    vTskH_AcctId  = Request("vTskH_AcctId")
    vTskH_Id      = Request("vTskH_Id")

    '...watch for next line - not sure if all conditions are met?
    If fNoValue(vTskH_AcctId) Then vTskH_AcctId = svCustAcctId

    '... before we even start, get member access info and the level 0 values
    sGetMemb (svMembNo)

    '...get customer reset status
    sGetCust svCustId

    '... check if there are any unlocking mechanisms in the Repository
    sUnlocks
  
    If Not fNoValue(vTskH_Id) Then 
      '...get level 0
      sGetTskH0 vTskH_AcctId, vTskH_Id
  
      '...ensure date/criteria/level OK
      If fTaskFilterOk Then
  
        '...password?
        If Not fNoValue(vTskH_Password) Then
          If Session("MyWorld2_PasswordEntered") <> vTskH_Password Then
            Session("MyWorld2_Password") = vTskH_Password
            Session("MyWorld2_Url") = "MyWorld2.asp?vTskH_Id=" & vTskH_Id & "&vTskH_No=" & vTskH_No
            Response.Redirect "TaskPassword.asp?" & Session("MyWorld2_Url")
          End If
        End If
  
  %>
  <div style="text-align: center">

    <!-- My Learning Grid is created here via MyWorld2.htm -->
    <%
      vBorder = 0 '...if we use collaborative tools set vBorder=1 to show border lines else vBorder = 0 then no border
      '...build grid and get the value of vBorder in MyWorld2.htm
      vGrid = fMyWorld2 (svCustAcctId, vTskH_Id, "live")
    %>

    <div style="text-align: center">
      <% If svMembLevel = 5 Then %>
      <a style="padding-right:40px;" href="MyWorld2.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vToggle=999"><!--[[-->Expand All Nodes<!--]]--></a> 
      <% End If %>

      <a href="MyWorld2.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vToggle=998"><!--[[-->Collapse All Nodes<!--]]--></a>

      <!-- show learners name and email address -->
      <%
      vInfo = fIf(Len(svMembFirstName)>0, svMembFirstName & " ", "") & fIf(Len(svMembLastName)>0, svMembLastName & " ", "") &  fIf(Len(svMembEmail)>0, "(" & svMembEmail & ")", "")
      If Len(Trim(vInfo)) > 0 Then Response.Write "<br><br>" & vInfo
      %>
    </div>


    <!--  Insert and display the main grid now, the first row is just to establish column widths -->
    <table style="width: 600px; margin: auto; border: 0;">
      <tr style="display:none">
        <td style="width: 25px"></td>
        <td style="width: 25px"></td>
        <td style="width: 550px">
      </tr>
      <% =vGrid %>
    </table>


    <!-- MyWorld2 Editors plus Administrators can edit task list -->
    <% If svMembLevel = 5 Or (svMembLevel = 4 And vMemb_MyWorld) Then %>
    <div style="text-align: center">
      <a href="MyWorld2Filters.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>">View Access Filters</a><%=f10%>
      <a href="TaskEdit2.asp?vTskH_Id=<%=vTskH_Id%>">Edit Task List</a>
    </div>
    <% End If %>
  </div>

  <%
      End If
  
    '...otherwise determine what Ids are available
    Else
      '...this defines the number of valid, level 0 filtered tasks

      vNoTasks = fNoTasks
  
      '...if there are no Ids then display message 
      If vNoTasks = 0 Then
  %>

  <div style="text-align: center">
    <h2>Sorry, but there are no Tasks available.</h2>
  </div>
  <% 
      '...if just one then rerun this page or if multiple tasks with different filters
      ElseIf vNoTasks = 1 Then 
        Response.Redirect "MyWorld2.asp?vTskH_AcctId=" & vTskH_AcctId & "&vTskH_Id=" & vTskH_Id
  
      '...else use must select from multiple Ids
      Else
  %>
  <form method="POST" action="MyWorld2.asp">
    <div style="text-align: center">
      <table class="table">
        <tr>
          <td style="text-align: center">
            <p style="text-align: center" class="c2">Select your task...</p>
            <p style="text-align: center">
              <select size="1" name="vTskH_Id"><%=fTaskOptions%></select>
              <input src="../Images/Buttons/Go_<%=svLang%>.gif" name="I1" type="image">
            </p>
          </td>
        </tr>
      </table>
    </div>
    <input type="hidden" name="vTskH_AcctId" value="<%=vTskH_AcctId%>">
  </form>
  <% 
      End If 
    End If  
  %>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
