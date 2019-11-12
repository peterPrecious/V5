<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<html>

<head>
  <title>ModuleDescription</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <%
    Server.Execute vShellHi

    '...display info if a modules is passed via a from
    vModId = Request.QueryString("vModId") 
    Dim vModId, vHours, aGoals, aSkillSet
    sGetMods (vModId)

  %>
  <table class="table">
    <tr>
      <td>
        <h1><%=vMods_Title & "</b>&nbsp;&nbsp;&nbsp;(" & vModId & ")"%></h1>
        <%=vMods_Desc%>

        <% If Len(vMods_Outline)>0 Then %>
        <h2><!--[[-->Module Outline<!--]]--></h2>
        <p class="c3"><%=vMods_Outline%></p>
        <% End If %>


        <% 
          If Len(vMods_Goals) > 0 Then 
            aGoals = Split (vMods_Goals,"::")
        %>
        <h3><!--[[-->Learning Objectives<!--]]--></h3>
        <h2><!--[[-->On completion of this<!--]]-->&nbsp;<%=vMods_Length%>&nbsp;<!--[[-->hour module you should be able to:<!--]]--></h2>
        <ul>
          <% For i = 0 to Ubound(aGoals) %>
          <li><%=aGoals(i)%></li>
          <% Next %> 
        </ul>
        <% End If %>


        <% 
          If Len(vMods_SkillSet) > 0 Then 
            aSkillSet = Split (vMods_SkillSet,"::")
        %>
        <h3>
          <!--[[-->Skill Set<!--]]--></h3>
        <h2></h2>
        <ul>
          <% For i = 0 to Ubound(aSkillSet) %>
          <li><%=aSkillSet(i)%></li>
          <% Next %>
        </ul>
        <% End If %>
      </td>
    </tr>
  </table>

  <% If Request("vClose") = "Y" Then '...if using a jWindow then mention closing... %>
  <h1><input onclick="javascript: window.close()" type="button" value="<%=fIf(svLang="FR", "Fermer", "Close")%>" name="bClose" class="button"></h1>
  <% End If %>


  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
