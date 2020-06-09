<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Skil.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
    Server.Execute vShellHi 
    
    Dim vFunction, vOrder
    vFunction = ""

    vOrder = fDefault(Request("vOrder"), "Id")


    '...update tables
    If Request("vFunction") = "add" Then
      sExtractJobs
      sInsertJobs
    ElseIf Request("vFunction") = "edit" Then
      sExtractJobs
      sUpdateJobs
    ElseIf Len(Request("vDelJobsId")) > 0 Then 
      vJobs_Id = Request("vDelJobsId")
      sDeleteJobs
    End If  
      
    If Len(Request("vForm")) = 0 Or Request("vFunction") = "del" Then

	%>
  <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vJobs_Id.value == "")
  {
    alert("Please enter a value for the \"Job Id (ie J1234EN)\" field.");
    theForm.vJobs_Id.focus();
    return (false);
  }

  if (theForm.vJobs_Id.value.length < 7)
  {
    alert("Please enter at least 7 characters in the \"Job Id (ie J1234EN)\" field.");
    theForm.vJobs_Id.focus();
    return (false);
  }

  if (theForm.vJobs_Id.value.length > 7)
  {
    alert("Please enter at most 7 characters in the \"Job Id (ie J1234EN)\" field.");
    theForm.vJobs_Id.focus();
    return (false);
  }

  var checkOK = "0123456789-J EN FR ES XX";
  var checkStr = theForm.vJobs_Id.value;
  var allValid = true;
  var validGroups = true;
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
  }
  if (!allValid)
  {
    alert("Please enter only digit and \"J EN FR ES XX\" characters in the \"Job Id (ie J1234EN)\" field.");
    theForm.vJobs_Id.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="JobsEdit.asp" target="_self" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
    <table border="0" width="100%" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td valign="bottom" colspan="2"><h1 align="center">The Jobs Table</h1><p>The Jobs Table contains a selection of learning programs assigned to categories. <b>Job Ids are unique to each account.</b> Click <b>Add</b> to create a new unique Job. Note, if possible the system will try to find a free Jobs Id either before or after the ones currently assigned.&nbsp; Jobs Id are formatted as J1234EN or J1234XX (for all languages).&nbsp; You can select an existing Jobs Id in the list below to edit or delete. If you try to creating a Job with an existing Jobs Id, the system will ignore the action.&nbsp; Note, once entered, it cannot be modified - only deleted and re-entered.</td>
      </tr>
      <tr>
        <td valign="bottom" align="left">Sort Table by Jobs
          <% If vOrder = "Title" Then %>
          <a class="c2" href="JobsEdit.asp?vOrder=Id">ID</a>
          <% Else %>
          <a class="c2" href="JobsEdit.asp?vOrder=Title">Title</a>
          <% End If %>
        </td>
        <th align="right" nowrap valign="bottom" class="c2">Add New Job with Id : <!--webbot bot="Validation" s-display-name="Job Id (ie J1234EN)" s-data-type="String" b-allow-digits="TRUE" s-allow-other-chars="J EN FR ES XX" b-value-required="TRUE" i-minimum-length="7" i-maximum-length="7" --><input type="text" name="vJobs_Id" size="6" maxlength="7" value="<%=fNextJobsId %>">&nbsp; <input type="submit" value="Add" name="bAdd" class="button">&nbsp; </th>
      </tr>
    </table>
    <input type="Hidden" name="vForm" value="Y">
  </form>
  <!---Edit List-->
  <table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse" height="26">
    <tr>
      <th align="left" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Jobs Id</th>
      <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Active</th>
<!--
      <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Level</th>
      <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Type</th>
-->
      <th align="left" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Jobs Title</th>
      <th align="left" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Program Title</th>
    </tr>
    <%
      '...read Jobs
      If vOrder = "Id" Then 
        sGetJobs_Rs
      Else
        sGetJobs_Rs_ByTitle      
      End If        
      Do While Not oRs3.Eof
       sReadJobs
    %>
    <tr>
      <td valign="top"><a href="JobsEdit.asp?vEditJobsId=<%=vJobs_Id%>&vForm=n"><%=vJobs_Id%></a>&nbsp; </td>
      <td valign="top" align="center"><%=fIf(vJobs_Active, "Y", "N")%></td>
<!--
      <td valign="top" align="center"><%=vJobs_Level%></td>
      <td valign="top" align="center"><%=vJobs_Type%></td>
-->

      <td valign="top"><%=vJobs_Title%></td>
      <td valign="top"><%=fJobMods(vJobs_Mods)%></td>
    </tr>
    <%  
        oRs3.MoveNext
      Loop
      Set oRs3 = Nothing
      sCloseDb3   
      
      Function fJobMods (vProgs)
        Dim aProgs, vProgId
        fJobMods = ""
        aProgs = Split(vProgs)
        For i = 0 To Ubound(aProgs)
          vProgId = aProgs(i)
          vProgId = "<a target='_blank' href='ProgramEdit.asp?vEditProgID=" & vProgId & "&vHidden=n&vLingo=ZZ&vRange=P100'>" & vProgId & "</a> "
          fJobMods = fJobMods & vProgId & "   " & fProgTitle(aProgs(i)) & "<br>"
        Next
      
      
      End Function 
      
 	 %>
    <tr>
      <td colspan="6" align="center">&nbsp;<h2><a href="CritEdit.asp">Group 1 Table</a>&nbsp; |&nbsp; <a href="SkilEdit.asp">Skills Table</a></h2></td>
    </tr>
  </table>
  <%
    Else

      If Len(Request.Form("vAddJobsId")) = 0 And Len(Request.Form("vForm")) > 0 Then 
        vJobs_Id = fNoQuote(Request.Form("vJobs_Id"))
        vFunction = "add"
      ElseIf Len(Request.QueryString("vEditJobsId")) > 0 Then 
        vJobs_Id = Request.QueryString("vEditJobsId")
        vFunction = "edit"
      Else
         Response.Redirect "JobsEdit.asp"          
      End If
  
      '...get the values (even if trying to add)
      If vJobs_Id <> "" Then sGetJobs vJobs_Id

      '...using skills?
      Dim aSkills, vSkillNo, aJobs_Ratings, aJobs_SkillMods

      If Len(fSkills) > 0 Then 
        aSkills   = Split(fSkills, "~")
        vSkillNo = Ubound(aSkills) + 1

        '...get skills ratings and mods and ensure same no as no of skills available
        '   there is an extra tilde between all ratings, hence the * 2 - 1 (except for the last value)
        '   initialize if more/less ratings on file then on skills table, or if first time

        If fNoValue(vJobs_Ratings) Or Len(vJobs_Ratings) <> vSkillNo * 2 - 1  Then   
          vJobs_Ratings   = Left("0~0~0~0~0~0~0~0~", vSkillNo * 2 - 1)
	        vJobs_SkillMods = Left("~~~~~~~", vSkillNo)
        End If

        '...put in array
        aJobs_Ratings   = Split(vJobs_Ratings, "~")
        aJobs_SkillMods = Split(vJobs_SkillMods, "~")

      Else
        vSkillNo = 0
      End If

  %>
  <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form2_Validator(theForm)
{

  var radioSelected = false;
  for (i = 0;  i < theForm.vJobs_Level.length;  i++)
  {
    if (theForm.vJobs_Level[i].checked)
        radioSelected = true;
  }
  if (!radioSelected)
  {
    alert("Please select one of the \"Job Level\" options.");
    return (false);
  }

  if (theForm.vJobs_Title.value == "")
  {
    alert("Please enter a value for the \"Job Title\" field.");
    theForm.vJobs_Title.focus();
    return (false);
  }

  if (theForm.vJobs_Title.value.length > 64)
  {
    alert("Please enter at most 64 characters in the \"Job Title\" field.");
    theForm.vJobs_Title.focus();
    return (false);
  }

  var radioSelected = false;
  for (i = 0;  i < theForm.vJobs_Type.length;  i++)
  {
    if (theForm.vJobs_Type[i].checked)
        radioSelected = true;
  }
  if (!radioSelected)
  {
    alert("Please select one of the \"Job Type\" options.");
    return (false);
  }

  if (theForm.vJobs_Mods.value == "")
  {
    alert("Please enter a value for the \"Job Programs\" field.");
    theForm.vJobs_Mods.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="JobsEdit.asp" target="_self" onsubmit="return FrontPage_Form2_Validator(this)" name="FrontPage_Form2" language="JavaScript">
    <input type="Hidden" name="vSkillNo" value="<%=vSkillNo%>"><input type="Hidden" name="vFunction" value="<%=vFunction%>"><input type="Hidden" name="vJobs_Id" value="<%=vJobs_Id%>">
    <table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td width="100%" valign="top" colspan="2"><h1 align="center">Edit the Jobs Table</h1><h2 align="left">The Jobs Table describes each Job Id and the associated learning programs (ie P1234EN...&nbsp; or P1234XX).&nbsp; If the Skills Table is used, then assign skills to this job and the rating of these skills.&nbsp; Only fill in as many skills as are applicable, when the system finds the first empty Skill, it assumes the list is complete.&nbsp; Click update when finished.</h2></td>
      </tr>
      <tr>
        <th align="right" width="20%" valign="top">Job Id :</th>
        <td valign="top"><h1><%=vJobs_Id%></h1></td>
      </tr>
      <tr>
        <th align="right" width="20%" valign="top">Job Level :</th>
        <td><!--webbot bot="Validation" s-display-name="Job Level" b-value-required="TRUE" --><input type="radio" value="A" name="vJobs_Level" <%=fcheck(vjobs_level, "a")%>>All<br><input type="radio" value="L" name="vJobs_Level" <%=fcheck(vjobs_level, "l")%>>Learners only (Level 2)<br><input type="radio" value="F" name="vJobs_Level" <%=fcheck(vjobs_level, "f")%>>Facilitators+ (Level 3+) </td>
      </tr>
      <tr>
        <th align="right" width="20%" valign="top">Job Title :</th>
        <td><!--webbot bot="Validation" s-display-name="Job Title" b-value-required="TRUE" i-maximum-length="64" --><input type="text" name="vJobs_Title" size="54" value="<%=vJobs_Title%>" maxlength="64"></td>
      </tr>
      <tr>
        <th align="right" width="20%" valign="top">Type :</th>
        <td><!--webbot bot="Validation" s-display-name="Job Type" b-value-required="TRUE" --><input type="radio" value="M" name="vJobs_Type" <%=fcheck(vjobs_type, "m")%>>Mandatory (core programs)<br><input type="radio" value="O" name="vJobs_Type" <%=fcheck(vjobs_type, "o")%>>Optional<br><input type="radio" value="A" name="vJobs_Type" <%=fcheck(vjobs_type, "a")%>>Assignable (typically via learner table)</td>
      </tr>
      <tr>
        <th align="right" width="20%" valign="top">Programs :<br>&nbsp;</th>
        <td><!--webbot bot="Validation" s-display-name="Job Programs" b-value-required="TRUE" --><textarea rows="3" name="vJobs_Mods" cols="51"><%=vJobs_Mods%></textarea><br>Enter a list of Programs that are of value to this job activity, separated by spaces, ie P1023EN P1034EN.&nbsp; You can also enter the Program Id as P1234XX P2345XX in which case the XX will be converted to the learner's language that was selected on signin.</td>
      </tr>
      <tr>
        <th align="right" width="20%" valign="top">Active:</th>
        <td><input type="radio" value="1" name="vJobs_Active" <%=fcheck(fsqlboolean(vjobs_active), 1)%>>Yes<br><input type="radio" value="0" name="vJobs_Active" <%=fcheck(fsqlboolean(vjobs_active), 0)%>>No (use this is building a new job string but before you want it reflected on reports)</td>
      </tr>
      <tr>
        <th align="right" valign="top" width="20%">Skills :</th>
        <td valign="top">Assign an expected rating for each skill that is required by this job plus appropriate learning programs, separated by spaces, ie P1002EN P0001EN, which will appear in My Learning if there is a skill gap between the learner&#39;s rating and the expected rating.&nbsp; A Skill Gap is determined after the learner completes the Skills form.<br>&nbsp;
        <table border="1" style="border-collapse: collapse" id="table1" cellpadding="3" bgcolor="#DDEEF9" bordercolor="#3977B6">
          <tr>
            <th nowrap>Skills<br>Required</th>
            <th nowrap>Expected<br>Rating (1-8)</th>
            <th nowrap>Programs<br>If Skills Gap</th>
          </tr>
          <tr>
            <th nowrap colspan="3" bgcolor="#FFFFFF">&nbsp;</th>
          </tr>
          <% For i = 1 To vSkillNo %>
          <tr>
            <td align="left"><p class="c1"><%=aSkills(i-1)%></p></td>
            <td align="center"><select size="1" name="vRate_<%=i%>"><%=fSkilRateOptions (aJobs_Ratings(i-1))%></select></td>
            <td align="center"><input type="text" name="vMods_<%=i%>" size="30" value="<%=aJobs_SkillMods(i-1)%>"></td>
          </tr>
          <% Next %>
        </table>
        </td>
      </tr>
      <tr>
        <td align="center" valign="top" colspan="2">
        
        &nbsp;<p>
        
        <input type="submit" value="Update" name="bUpdate" class="button"></p></p><h2>Ensure you do <b>NOT </b>delete any Jobs Id that may be used in the Group 1 Table ... 
        <input type="button" onclick="jconfirm('JobsEdit.asp?vDelJobsId=<%=vJobs_Id%>&vFunction=del', 'Ok to delete?')" value="Delete" name="bDelete" class="button">
        
        
        
        <h2><a href="JobsEdit.asp">Jobs List</a>&nbsp; |&nbsp; <a href="CritEdit.asp">Group 1 Table</a>&nbsp; |&nbsp; <a href="SkilEdit.asp">Skills Table</a></h2></td>
      </tr>
    </table>
  </form>
  <%
    End If
    
    Server.Execute vShellLo 
  %>

</body>

</html>
