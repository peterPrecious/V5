<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<% 
  Dim vJobsNo, vJobsTitle, vMods, aMods

  sGetJobsByMemb       '...get the job categories that apply to this criteria / member
  sGetMemb svMembNo    '...get the user record to find job if any previous title selection

  If Request("vForm") = "y" Then
    '...update new mods/progs then go to My Learning (change any commas to spaces)
    sUpdateMembPrograms svMembNo, Replace(Request.Form("vMembPrograms"), ", ", " ")
    Response.Redirect "MyWorld.asp?vTskH_Id=" & Request("vTskH_Id")
  End If

  Select Case vMemb_JobsNo
    Case 0 : vJobsTitle = vJobs_Title0 : vMods = vJobs_Mods0
    Case 1 : vJobsTitle = vJobs_Title1 : vMods = vJobs_Mods1
    Case 2 : vJobsTitle = vJobs_Title2 : vMods = vJobs_Mods2
    Case 3 : vJobsTitle = vJobs_Title3 : vMods = vJobs_Mods3
    Case 4 : vJobsTitle = vJobs_Title4 : vMods = vJobs_Mods4
    Case 5 : vJobsTitle = vJobs_Title5 : vMods = vJobs_Mods5
    Case 6 : vJobsTitle = vJobs_Title6 : vMods = vJobs_Mods6
    Case 7 : vJobsTitle = vJobs_Title7 : vMods = vJobs_Mods7
    Case 8 : vJobsTitle = vJobs_Title8 : vMods = vJobs_Mods8
  End Select
%>
<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="1" style="border-collapse: collapse" cellspacing="0" cellpadding="4" bordercolor="#DDEEF9" width="100%" id="table1">
    <form method="POST" action="JobSkills.asp">
      <input type="hidden" name="vForm" value="y"><input type="hidden" name="vTskH_Id" value="<%=Request("vTskH_Id")%>">
      <tr>
        <td align="left" colspan="4">
        <h1 align="center"><br>My Learning Assessment - Step 2 (of 2)</h1>
        <% If Len(Trim(vMemb_Programs)) > 0 Then %>
        <h6 align="center"><%=svMembFirstName%>, since you have already created your Training Plan, new module selections will overwrite any previously selected learning modules.&nbsp; If you do not wish to change any selections, click on the globe to return to My Learning.</h6>
        <% End If %>
        <p>Here is a list of Learning Modules relevant to your specific job title.</p>
        <ol>
          <li>Click on the Title to preview the module.</li>
          <li>Click on [Description] to view a course description.</li>
          <li>Click on [Take Test] to test your current knowledge in this subject area.&nbsp; If you do NOT pass the test, we strongly suggest that you include the module in your training plan.&nbsp; Test results are for your personal information only.&nbsp; Results will NOT be recorded or saved.</li>
          <li>Click &quot;Yes&quot; beside each module you want to include in your training plan then click &quot;update&quot;.&nbsp; Selected modules will appear in &quot;My Learning&quot; under &quot;My Training Plan&quot;.</li>
        </ol>
        </td>
      </tr>
      <tr>
        <th align="left" bgcolor="#DDEEF9" colspan="2">Learning Modules</th>
        <th align="left" bgcolor="#DDEEF9" rowspan="2">&nbsp;</th>
        <th bgcolor="#DDEEF9" rowspan="2">Include in My Training Plan?</th>
      </tr>
      <tr>
        <th align="left" bgcolor="#DDEEF9">Id</th>
        <th align="left" bgcolor="#DDEEF9">Title</th>
      </tr>
      <%
        '...process Mods
        aMods = Split(vMods, " ")
        For i = 0 To Ubound(aMods)
	    %>
      <tr>
        <td valign="top" align="left"><%=aMods(i)%></td>
        <td valign="top" align="left"><a href="javascript:zmodulewindow('<%=aMods(i)%>')"><%=fModsTitle(aMods(i))%></a></td>
        <td valign="top" nowrap>[<a href="javascript:SiteWindow('ModuleDescription.asp?vClose=Y&vModId=<%=aMods(i)%>')">Description</a>] [<a href="javascript:SiteWindow('Test.asp?vClose=Y&vModId=<%=aMods(i)%>')">Take Test</a>]</td>
        <td valign="top" align="center"><input type="checkbox" name="vMembPrograms" value="<%=aMods(i)%>" <%=fif(instr(vmemb_programs, amods(i)) > 0,"checked", "")%>>Yes</td>
      </tr>
      <%  
        Next
      %>
      <tr>
        <td valign="top" align="center" colspan="4">&nbsp;<p><a href="MyWorld.asp?vTskH_Id=<%=Request("vTskH_Id")%>"><img border="0" src="../Images/Icons/World.gif" alt="<%=Server.HtmlEncode("<!--[[-->Return to My Learning<!--]]-->")%>"></a></p>
        <p><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I4" type="image"><br>&nbsp; </p>
        </td>
      </tr>
    </form>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
