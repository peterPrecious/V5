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
<!--#include file = "MyWorldCode.asp"-->
<!--#include file = "MyWorldCodeRoutines.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script src="/V5/Inc/Launch.js"></script>
  <script src="/V5/Inc/LaunchObjects.js"></script>
  <script>
    function fAlert() {
      var vPhrase = "<%=fPhraH(000366)%>"
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
        var vWs = WebService("MyWorldCode_ws.asp", vParams)
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
  <title>My Learning</title>
</head>

<body>

  <% 
    Server.Execute vShellHi 

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
          If Session("MyWorld_PasswordEntered") <> vTskH_Password Then
            Session("MyWorld_Password") = vTskH_Password
            Session("MyWorld_Url") = "MyWorld.asp?vTskH_Id=" & vTskH_Id & "&vTskH_No=" & vTskH_No
            Response.Redirect "TaskPassword.asp?" & Session("MyWorld_Url")
          End If
        End If
  
  %>
  <div style="text-align: center">

    <!-- My Learning Grid is created here via MyWorld.htm -->
    <%
      vBorder = 0 '...if we use collaborative tools set vBorder=1 to show border lines else vBorder = 0 then no border
      '...build grid and get the value of vBorder in MyWorld.htm
      vGrid = fMyWorld (svCustAcctId, vTskH_Id, "live")
    %>


    <div style="text-align: center">
      <% If svMembLevel = 5 Then %>
      <a href="MyWorld.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vToggle=999"><!--webbot bot='PurpleText' PREVIEW='Expand All Nodes'--><%=fPhra(000135)%></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
      <% End If %>

      <a href="MyWorld.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>&vToggle=998"><!--webbot bot='PurpleText' PREVIEW='Collapse All Nodes'--><%=fPhra(000105)%></a>

      <!-- show learners name and email address -->
      <%
      vInfo = fIf(Len(svMembFirstName)>0, svMembFirstName & " ", "") & fIf(Len(svMembLastName)>0, svMembLastName & " ", "") &  fIf(Len(svMembEmail)>0, "(" & svMembEmail & ")", "")
      If Len(Trim(vInfo)) > 0 Then Response.Write "<br><br>" & vInfo
      %>   
    </div>


    <!--  Display the main grid now -->
    <table style="width:600px;margin:auto;">
      <% =vGrid %>
    </table>


    <!-- MyWorld Editors plus Administrators can edit task list -->
    <% If svMembLevel = 5 Or (svMembLevel = 4 And vMemb_MyWorld) Then %>
      <div style="text-align: center">
        <a href="MyWorldFilters.asp?vTskH_AcctId=<%=vTskH_AcctId%>&vTskH_Id=<%=vTskH_Id%>">View Access Filters</a><%=f10%>
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
        Response.Redirect "MyWorld.asp?vTskH_AcctId=" & vTskH_AcctId & "&vTskH_Id=" & vTskH_Id
  
      '...else use must select from multiple Ids
      Else
  %>
  <form method="POST" action="MyWorld.asp">
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

  <style>
    td { text-align: left !important; }
  </style>

</body>

</html>


