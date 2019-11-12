<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_SSet.asp"-->
<%
  Dim vL, vG
  vL = fDefault(Request("vL"), "EN")
  vG = fDefault(Request("vG"), "*")

  If Request("vRefresh") = "y" Then sInitSset
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
      <td colspan="4">
      <h3>Skill Set Table</h3>
      <p>This lists all the skills by module group and language.&nbsp; To modify a skill set, click on the module in question, change the skill set the close the module window.&nbsp; When finished, you can update the Skill Set table with new skill sets by clicking on the update Skill Set Table link at the bottom of this page.&nbsp; Note: do this INFREQUENTLY as it consumes considerable resources.</p>
      <form method="POST" action="ModSkillSet.asp">
        <div align="center">
          <table border="1" id="table1" cellpadding="3" style="border-collapse: collapse" bordercolor="#DCEDF8">
            <tr>
              <th nowrap align="right" valign="top">Select Languages :</th>
              <td>
                 <input type="radio" value="EN" name="vL" <%=fCheck(vl, "EN")%>>EN&nbsp;&nbsp; 
                 <input type="radio" value="FR" name="vL" <%=fCheck(vl, "FR")%>>FR&nbsp;&nbsp; 
                 <input type="radio" value="ES" name="vL" <%=fCheck(vl, "ES")%>>ES&nbsp;&nbsp; 
                 <input type="radio" value="PT" name="vL" <%=fCheck(vl, "PT")%>>PT</td>
            </tr>
            <tr>
              <th nowrap align="right" valign="top">Display/Select by&nbsp;&nbsp; <br>Module Group :</th>
              <td>
                <input type="radio" value="X" name="vG" <%=fCheck(vg, "X")%>>Do not break down into module groups (a group is defined by module&#39;s first digit)<br>
                <input type="radio" value="*" name="vG" <%=fCheck(vg, "*")%>>Show All Groups, or<br>&nbsp;&nbsp;&nbsp;&nbsp; show modules starting with...<br> 
                <input type="radio" value="0" name="vG" <%=fCheck(vg, "0")%>>0xxx&nbsp;&nbsp; 
                <input type="radio" value="1" name="vG" <%=fCheck(vg, "1")%>>1xxx&nbsp;&nbsp; 
                <input type="radio" value="2" name="vG" <%=fCheck(vg, "2")%>>2xxx&nbsp;&nbsp; 
                <input type="radio" value="3" name="vG" <%=fCheck(vg, "3")%>>3xxx&nbsp;&nbsp; 
                <input type="radio" value="4" name="vG" <%=fCheck(vg, "4")%>>4xxx&nbsp;&nbsp; 
                <input type="radio" value="5" name="vG" <%=fCheck(vg, "5")%>>5xxx&nbsp;&nbsp; 
                <input type="radio" value="6" name="vG" <%=fCheck(vg, "6")%>>6xxx&nbsp;&nbsp; 
                <input type="radio" value="7" name="vG" <%=fCheck(vg, "7")%>>7xxx&nbsp;&nbsp; 
                <input type="radio" value="8" name="vG" <%=fCheck(vg, "8")%>>8xxx&nbsp;&nbsp; 
                <input type="radio" value="9" name="vG" <%=fCheck(vg, "9")%>>9xxx&nbsp;&nbsp; 
              </td>
            </tr>
            <tr>
              <th nowrap align="right">&nbsp;</th>
              <td> <input border="0" src="../Images/Buttons/Go_<%=svLang%>.gif" name="I1" type="image"></td>
            </tr>
          </table>
        </div>
        <p align="center">: </p>
      </form>
      <p>&nbsp;</p></td>
    </tr>
    <tr>
      <td align="center"><h1>Module<br>Group</h1></td>
      <td align="center"><h1>Lang</h1></td>
      <td align="left"><h1>Skill Set</h1></td>
      <td align="left"><h1>Modules</h1></td>
    </tr>
    <%
      '...read Prog
      Dim aModId, vModId
      sGetSset_Rs vL, vG
      Do While Not oRsBase.Eof 
        sReadSset
    %>
    <tr>
      <td valign="top" align="center"><%=vSset_Group%></td>
      <td valign="top" align="center"><%=vSset_Lang%></td>
      <td valign="top"><%=vSset_Id%></td>
      <td valign="top">
      <%
        aModId = Split(vSset_ModIds, " ")
        For i = 0 To Ubound(aModId)
      %> 
      <a target="_blank" href="ModuleEdit.asp?vEditModsId=<%=aModId(i)%>&vHidden=n"><%=aModId(i)%></a> 
      <% 
        Next
      %> 
      </td>
    </tr>
    <%  
        oRsBase.MoveNext
      Loop
      Set oRsBase = Nothing
      sCloseDbBase    
  	%>
    <tr>
      <td colspan="4" align="center"><p><br>&nbsp;</p><p>If absolutely necessary: <a href="ModSkillSet.asp?vRefresh=y">Click here to update the Skill Set Table</a></p><p><img border="0" src="../Images/Icons/Bang.gif" width="18" height="22">Note: clicking on the above link will recreate the table and requires a several minutes.&nbsp; <br>Nothing will appear until the update is finished and the refreshed table will be redisplayed.<br>&nbsp;</p></td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
