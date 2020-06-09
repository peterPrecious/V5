<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Skil.asp"-->

<%
   Dim vSkillNo, vSkill_1, vSkill_2, vSkill_3, vSkill_4, vSkill_5, aMemb_Skills, vMsg, vMembNo

   vMembNo = svMembNo

   If Len(Request.Form("vForm")) > 0 Then

     vMembNo = Request.Form("vMembNo")

     '...figure out how many skills to check
     vSkillNo = Request.Form("vSkillNo")

     If vSkillNo >= 1 Then If Len(Request.Form("1")) > 0 Then vMemb_Skills = vMemb_Skills & Request.Form("1") & "~" Else vMsg = "Please assess all skills"
     If vSkillNo >= 2 Then If Len(Request.Form("2")) > 0 Then vMemb_Skills = vMemb_Skills & Request.Form("2") & "~" Else vMsg = "Please assess all skills"
     If vSkillNo >= 3 Then If Len(Request.Form("3")) > 0 Then vMemb_Skills = vMemb_Skills & Request.Form("3") & "~" Else vMsg = "Please assess all skills"
     If vSkillNo >= 4 Then If Len(Request.Form("4")) > 0 Then vMemb_Skills = vMemb_Skills & Request.Form("4") & "~" Else vMsg = "Please assess all skills"
     If vSkillNo >= 5 Then If Len(Request.Form("5")) > 0 Then vMemb_Skills = vMemb_Skills & Request.Form("5") & "~" Else vMsg = "Please assess all skills"
     If vSkillNo >= 6 Then If Len(Request.Form("6")) > 0 Then vMemb_Skills = vMemb_Skills & Request.Form("6") & "~" Else vMsg = "Please assess all skills"
     If vSkillNo >= 7 Then If Len(Request.Form("7")) > 0 Then vMemb_Skills = vMemb_Skills & Request.Form("7") & "~" Else vMsg = "Please assess all skills"
     If vSkillNo >= 8 Then If Len(Request.Form("8")) > 0 Then vMemb_Skills = vMemb_Skills & Request.Form("8") & "~" Else vMsg = "Please assess all skills"

     If Len(vMsg) = 0 Then    
       '...strip off trialing "~"     
       vMemb_Skills = Left(vMemb_Skills, Len(vMemb_Skills)-1)
       '...update member record
       sUpdateMembSkills vMembNo, vMemb_Skills
       vMsg = "Updated Successfully !"
     End If

   Else

     '...Get the member skills
     sGetMemb vMembNo

     If fNoValue(vMemb_Skills) Then 
       vMemb_Skills = "0~0~0~0~0~0~0~0"
     End If
     
   End If
%>

<html>
  <head>
    <meta charset="UTF-8">
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  </head>

  <body>

  <% Server.Execute vShellHi %>

  <table border="1" style="border-collapse: collapse" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" width="100%">
    <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vMembNo.selectedIndex < 0)
  {
    alert("Please select one of the \"User\" options.");
    theForm.vMembNo.focus();
    return (false);
  }

  if (theForm.vMembNo.selectedIndex == 0)
  {
    alert("The first \"User\" option is not a valid selection.  Please choose one of the other options.");
    theForm.vMembNo.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="SkillsAssessment.asp" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
      <input type="hidden" name="vForm" value="y"><input type="hidden" name="vTskH_Id" value="<%=Request("vTskH_Id")%>">
      <tr>
        <td align="center">
        <h1><br>Skills Assessment</h1>


        <% If Len(vMsg) = 0 Then %>
        In order to assess job skills, please select the appropriate rating then click &quot;update&quot;.<br>Clicking on Skill Gap Analysis will display how these skills compare to expected skills.<h2>Note: for a description of the skills or ratings, hold your mouse over the appropriate title.<br>&nbsp;</h2>
        <% Else %>
        <h3>Updated Successfully !</h3>
        <% End If %>


        <div align="center">
          <table border="1" style="border-collapse: collapse" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9">
            <tr>
              <td colspan="6" align="center"><br><span style="background-color: #FFFF00">[Developer&#39;s note: This dropdown menu only appears for facilitators <br>(ie department heads) who wish to assess learners (ie their staff).]</span><br><!--webbot bot="Validation" s-display-name="User" b-value-required="TRUE" b-disallow-first-item="TRUE" --><select size="1" name="vMembNo">
              <option>Select Learner</option>
							<%=fMembDropdown (vMembNo)%>
              </select><br>&nbsp;</td>
            </tr>
            <tr>
              <th valign="top" rowspan="2" width="125">Skills<br><br><img border="0" src="../Images/Icons/ArrowDown.gif" width="18" height="22"></th>
              <th colspan="5" align="left" width="125" height="30">Rating Scale&nbsp; <img border="0" src="../Images/Icons/ArrowRight.gif" align="absmiddle"></th>
            </tr>
            <tr>
              <td valign="top" align="center" class="c1" width="125">1 - <a title="Skills are insufficient to support role or task." href="javascript:;">Challenges to Meet</a>&nbsp; </td>
              <td valign="top" align="center" class="c1" width="125">2 - <a title="Skills developing in this area.  Supervision and direction is required." href="javascript:;">Developing</a>&nbsp; </td>
              <td valign="top" align="center" class="c1" width="125">3 - <a title="Skills are adequate and consistent to support role.  Occasional supervision and direction is required." href="javascript:;">Competent</a>&nbsp; </td>
              <td valign="top" align="center" class="c1" width="125">4 - <a title="Highly competent. Can be viewed as a resource that can deliver effort.  Provides input and advice in area of expertise to other colleagues." href="javascript:;">Proficient</a>&nbsp; </td>
              <td valign="top" align="center" class="c1" width="125">5 - <a title="Highly skilled.  Acts as a role model for others within team. Understands the importance of transferring knowledge to colleagues.  Views as the &quot;Go To&quot; resources within the team." href="javascript:;">Expert</a>&nbsp; </td>
            </tr>
            <% 
               vSkillNo = 0
               aMemb_Skills = Split(vMemb_Skills, "~")

               '...display all skills
               sGetSkil_Rs
                 Do While Not oRs4.Eof
                 sReadSkil
                 vSkillNo = vSkillNo + 1
                
            %>
            <tr>
              <td align="left" width="125"><p class="c1"><a title="<%=vSkil_Desc%>" href="javascript:;"><%=vSkil_Id%></a></p></td>
              <td align="center" width="125"><input type="radio" value="1" name="<%=vSkillNo%>" <%=fcheck(amemb_skills(vskillno-1), "1")%>></td>
              <td align="center" width="125"><input type="radio" value="2" name="<%=vSkillNo%>" <%=fcheck(amemb_skills(vskillno-1), "2")%>></td>
              <td align="center" width="125"><input type="radio" value="3" name="<%=vSkillNo%>" <%=fcheck(amemb_skills(vskillno-1), "3")%>></td>
              <td align="center" width="125"><input type="radio" value="4" name="<%=vSkillNo%>" <%=fcheck(amemb_skills(vskillno-1), "4")%>></td>
              <td align="center" width="125"><input type="radio" value="5" name="<%=vSkillNo%>" <%=fcheck(amemb_skills(vskillno-1), "5")%>></td>
            </tr>
            <%  
                 oRs4.MoveNext
               Loop
               Set oRs4 = Nothing
               sCloseDb4   
            %> 
            <input type="hidden" name="vSkillNo" value="<%=vSkillNo%>">
          </table>

          <% If Len(Request("vTskH_Id")) > 0 Then %> <br><br><a href="MyWorld.asp?vTskH_Id=<%=Request("vTskH_Id")%>"><img border="0" src="../Images/Icons/World.gif" alt="<%=Server.HtmlEncode("<!--webbot bot='PurpleText' PREVIEW='Return to My Learning'--><%=fPhra(000331)%>")%>"></a> <% End If %> <br><br><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I4" type="image"><br><br><br><a href="SkillsGapAnalysis.asp">Skill Gap Analysis </a><br>&nbsp;</div>
        </td>
      </tr>
    </form>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>
</html>




