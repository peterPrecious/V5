<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Skil.asp"-->

<%
   Dim vMembNo, vSkillNo
   vMembNo = svMembNo
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
  <table border="1" style="border-collapse: collapse" width="100%" id="table1" bordercolor="#DDEEF9">
    <tr>
      <td align="center">
      <h1><br>Skill Gap Analysis Report</h1>
      <table border="0" style="border-collapse: collapse" cellpadding="0" width="400">
        <tr>
          <td>
          <div align="left">
            <table border="1" style="border-collapse: collapse" id="table5" cellpadding="3" cellspacing="3" bordercolor="#DDEEF9">
              <tr>
                <td colspan="2">
                <h1>Select View</h1>
                </td>
              </tr>
              <tr>
                <td><!--webbot bot="Validation" s-display-name="User" b-value-required="TRUE" b-disallow-first-item="TRUE" --><select size="1" name="vMembNo">
                <option>Select Learner</option>
                <%=fMembDropdown (vMembNo)%></select></td>
                <td bordercolor="#FFFFFF"><input border="0" src="../Images/Buttons/Go_<%=svLang%>.gif" name="I4" type="image" width="36" height="19"></td>
              </tr>
            </table>
          </div>
          </td>
          <td></td>
          <td>
          <div align="right">
            <table border="1" style="border-collapse: collapse" id="table6" cellpadding="3" cellspacing="0" bordercolor="#DDEEF9">
              <tr>
                <td colspan="2"><p class="c1">Legend:</p></td>
              </tr>
              <tr>
                <td>Employer Rating</td>
                <td>
                <table border="0" style="border-collapse: collapse" width="100%" id="table8" cellpadding="0">
                  <tr>
                    <td bgcolor="#FFFF00" height="10" width="15"></td>
                  </tr>
                </table>
                </td>
              </tr>
              <tr>
                <td>Job Requirement</td>
                <td>
                <table border="0" style="border-collapse: collapse" width="100%" id="table10" cellpadding="0">
                  <tr>
                    <td bgcolor="#008000" height="10" width="15"></td>
                  </tr>
                </table>
                </td>
              </tr>
              <tr>
                <td>Employee Rating</td>
                <td>
                <table border="0" style="border-collapse: collapse" width="100%" id="table9" cellpadding="0">
                  <tr>
                    <td bgcolor="#0000FF" height="10" width="15"></td>
                  </tr>
                </table>
                </td>
              </tr>
            </table>
          </div>
          </td>
        </tr>
      </table>
      <br>&nbsp;
      <div align="center">
        <table border="1" cellspacing="0" bordercolor="#DDEEF9" cellpadding="3" style="border-collapse: collapse" id="table2" width="400">
          <tr>
            <td rowspan="2">
            <h1>Skills </h1>
            </td>
            <td colspan="5" align="center"><p class="c1">Scores</p></td>
            <td colspan="2" align="center"><p class="c1">Gap Analysis</p></td>
          </tr>
          <tr>
            <td class="c2">1</td>
            <td class="c2">2</td>
            <td class="c2">3</td>
            <td class="c2">4</td>
            <td class="c2">5</td>
            <td align="center" class="c2">Employee</td>
            <td align="center" class="c2">Manager</td>
          </tr>
          <%
            '...display all skills
            sGetSkil_Rs
              Do While Not oRs4.Eof
              sReadSkil
              vSkillNo = vSkillNo + 1
          %>
          <tr>
            <td colspan="8">&nbsp;</td>
          </tr>
          <tr>
            <td rowspan="3"><%=vSkil_Id%></td>
            <td bgcolor="#FFFF00" height="10"></td>
            <td bgcolor="#FFFF00" height="10"></td>
            <td bgcolor="#FFFF00" height="10"></td>
            <td height="10"></td>
            <td height="10"></td>
            <td rowspan="3" align="center" class="c3">0</td>
            <td rowspan="3" align="center" class="c3">&nbsp;</td>
          </tr>
          <tr>
            <td bgcolor="#008000" height="10"></td>
            <td bgcolor="#008000" height="10"></td>
            <td bgcolor="#008000" height="10"></td>
            <td height="10"></td>
            <td height="10"></td>
          </tr>
          <tr>
            <td bgcolor="#0000FF" height="10"></td>
            <td bgcolor="#0000FF" height="10"></td>
            <td bgcolor="#0000FF" height="10"></td>
            <td height="10"></td>
            <td height="10"></td>
          </tr>
          <%  
              oRs4.MoveNext
            Loop
            Set oRs4 = Nothing
            sCloseDb4   
          %>
          <tr>
            <td colspan="8">&nbsp;</td>
          </tr>
        </table>
        <p><br><a href="MyWorld.asp?vTskH_Id=<%=Request("vTskH_Id")%>"><img border="0" src="../Images/Icons/World.gif" alt="<%=Server.HtmlEncode("<!--[[-->Return to My Learning<!--]]-->")%>"></a></p><p><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a></p></div>
      <p></p></td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
