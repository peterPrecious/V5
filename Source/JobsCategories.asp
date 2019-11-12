<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->

<% 
  Dim vJobsNo

  sGetJobsByMemb       '...get the job categories that apply to this criteria / member
  sGetMemb svMembNo    '...get the user record to find job if any previous title selection

  If Request("vForm") = "y" Then
    '...If they have selected a new job title, then update Memb with new JobsNo and kill any old Programs
    If Request.Form("vJobsNo") <> vMemb_JobsNo Then 
      sUpdateMembJobsNo svMembNo, Request.Form("vJobsNo")
    End If
    Response.Redirect "JobSkills.asp?vTskH_Id=" & Request("vTskH_Id")
  End If
  
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
  
  <table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <td align="center">
      <h1><br>My Learning Assessment - Step 1 (of 2)</h1>

      <% If Len(Trim(vMemb_Programs)) > 0 Then %> 
      <h6 align="left"><%=svMembFirstName%>, since you have already created your Training Plan, changing Job Titles will eliminate any previously selected learning modules.&nbsp; If you do not wish to change your title, click &quot;continue&quot; to review you previously selected learning modules. Otherwise, select a new Job Title and click &quot;update to view a list of learning modules relevant to your job. <br>&nbsp;</h6>
      <% Else %>
      <p>Select the Job Title that best describes your responsibilities then <br>click &quot;update&quot; to view a list of relevant learning modules.</p>
      <% End If %>
      
      <div align="center">
        <form method="POST" action="JobsCategories.asp">
 
          <input type="hidden" name="vForm" value="y">
          <input type="hidden" name="vTskH_Id" value="<%=Request("vTskH_Id")%>">
 

          <table border="1" style="border-collapse: collapse" bordercolor="#DDEEF9" cellspacing="0" cellpadding="3" width="50%">

            <% 
              If vJobs_Eof Then Response.Redirect "Error.asp?vErr=" & Server.UrlEncode("No Jobs have been defined for your category of work.")
            %>

            <tr>
              <th nowrap align="left" bgcolor="#DDEEF9" height="30">Job Titles</th>
              <th nowrap bgcolor="#DDEEF9" height="30">Select</th>
            </tr>


            <tr>
              <td><h1><br><%=vJobs_Title0%></h1></td>
              <td align="center"><br><input type="radio" value="0" name="vJobsNo" <%=fCheck("0", vMemb_JobsNo)%>></td>
            </tr>
            <% If Len(vJobs_Title1) > 0 Then %>
            <tr>
              <td>&nbsp;&nbsp;&nbsp; <%=vJobs_Title1%></td>
              <td align="center"><input type="radio" value="1" name="vJobsNo" <%=fCheck("1", vMemb_JobsNo)%>></td>
            </tr>
            <% End If %> <% If Len(vJobs_Title2) > 0 Then %>
            <tr>
              <td>&nbsp;&nbsp;&nbsp; <%=vJobs_Title2%></td>
              <td align="center"><input type="radio" value="2" name="vJobsNo" <%=fCheck("2", vMemb_JobsNo)%>></td>
            </tr>
            <% End If %> <% If Len(vJobs_Title3) > 0 Then %>
            <tr>
              <td>&nbsp;&nbsp;&nbsp; <%=vJobs_Title3%></td>
              <td align="center"><input type="radio" value="3" name="vJobsNo" <%=fCheck("3", vMemb_JobsNo)%>></td>
            </tr>
            <% End If %> <% If Len(vJobs_Title4) > 0 Then %>
            <tr>
              <td>&nbsp;&nbsp;&nbsp; <%=vJobs_Title4%></td>
              <td align="center"><input type="radio" value="4" name="vJobsNo" <%=fCheck("4", vMemb_JobsNo)%>></td>
            </tr>
            <% End If %> <% If Len(vJobs_Title5) > 0 Then %>
            <tr>
              <td>&nbsp;&nbsp;&nbsp; <%=vJobs_Title5%></td>
              <td align="center"><input type="radio" value="5" name="vJobsNo" <%=fCheck("5", vMemb_JobsNo)%>></td>
            </tr>
            <% End If %> <% If Len(vJobs_Title6) > 0 Then %>
            <tr>
              <td>&nbsp;&nbsp;&nbsp; <%=vJobs_Title6%></td>
              <td align="center"><input type="radio" value="6" name="Job <%=fCheck("6", vMemb_JobsNo)%>"></td>
            </tr>
            <% End If %> <% If Len(vJobs_Title7) > 0 Then %>
            <tr>
              <td>&nbsp;&nbsp;&nbsp; <%=vJobs_Title7%></td>
              <td align="center"><input type="radio" value="7" name="vJobsNo" <%=fCheck("7", vMemb_JobsNo)%>></td>
            </tr>
            <% End If %> <% If Len(vJobs_Title8) > 0 Then %>
            <tr>
              <td>&nbsp;&nbsp;&nbsp; <%=vJobs_Title8%></td>
              <td align="center"><input type="radio" value="8" name="vJobsNo" <%=fCheck("8", vMemb_JobsNo)%>></td>
            </tr>
            <% End If %>
          </table>

          <p><a href="MyWorld.asp?vTskH_Id=<%=Request("vTskH_Id")%>"><img border="0" src="../Images/Icons/World.gif"></a></p>
          <p><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I4" type="image">
         
          <% If Len(Trim(vMemb_Programs)) > 0 Then %> 
          &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="JobSkills.asp"><img border="0" src="../Images/Buttons/Continue_<%=svLang%>.gif" alt="Continue without updating..."></a>
          <% End If %>
          
          
          <br><br><br><a href="JobsEdit.asp">Edit Job Categories</a><br>&nbsp;</p>
        </form>
      </div>
      </td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>
</html>
