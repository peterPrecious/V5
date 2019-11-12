<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<% 
    Session("Ecom_Mods") = Request("vModsId")
    Dim vHours, aGoals, aSkillSet, vReturnToPrograms
    sGetMods Session("Ecom_Mods")

    vReturnToPrograms = fIf(Session("Ecom_Media") = "Group2", "Ecom3Programs.asp", "Ecom2Programs.asp")
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Module Description</title>
</head>

<body>

  <%
    Server.Execute vShellHi
  %>
  <table style="width: 100%">
    <tr>
      <td><img border="0" src="../Images/Ecom/Module.gif" width="74" height="64">&nbsp; </td>

      <td align="center"><h1><!--webbot bot='PurpleText' PREVIEW='Module Description'--><%=fPhra(000176)%></h1><p class="c2"> <!--webbot bot='PurpleText' PREVIEW='The outline, objectives and skill set are detailed here for this module.'--><%=fPhra(000255)%></p>
      <% If Session("Ecom_Media") = "" Then '...ie My Content rather then More Content %>
        <h2 align="center"> <a <%=fStatX%> href="javascript:history.back(1)"><!--webbot bot='PurpleText' PREVIEW='Return to Modules'--><%=fPhra(000227)%></a><br>&nbsp;</h2>
      <% Else %>
        <h2 align="center"><!--webbot bot='PurpleText' PREVIEW='Return to'--><%=fPhra(000226)%> <a <%=fStatX%> href="<%=vReturnToPrograms%>".asp?vCatlId=<%=Session("Ecom_Catl")%>"><!--webbot bot='PurpleText' PREVIEW='Programs'--><%=fPhra(000203)%></a> | <a <%=fStatX%> href="javascript:history.back(1)">Modules</a></h2>
      <% End If %>
        <br />
      </td>
    </tr>
  </table>

  <table style="width: 100%">
    <tr>
      <td bgcolor="#FFFFFF" align="left" valign="top">
      <p align="center"><span class="c1"><%=vMods_Title%></span>&nbsp;&nbsp;&nbsp;<span class="c2">(<%=vMods_Id%>)</span></p>
      
      <h2 align="center"><%=vMods_Desc%></h2>
 
      <% If Len(vMods_Outline)>0 Then %>
      <h2>
      <!--webbot bot='PurpleText' PREVIEW='Module Outline'--><%=fPhra(000177)%></h2>
      <p><%=vMods_Outline%></p>
      <% 
        End If 
        
        If Len(vMods_Goals) > 0 Then 
          aGoals = Split (vMods_Goals,"::")
      %>
      <h2>
        <!--webbot bot='PurpleText' PREVIEW='Learning Objectives'--><%=fPhra(000169)%></h2>
      <p>
      <!--webbot bot='PurpleText' PREVIEW='On completion of this'--><%=fPhra(000417)%>&nbsp;<%=vMods_Length%>&nbsp;
      <!--webbot bot='PurpleText' PREVIEW='hour module you should be able to:'--><%=fPhra(000418)%></p>
      <ul>
	      <% For i = 0 to Ubound(aGoals) %>
        <li><%=aGoals(i)%></li> 
	      <% Next %>
      </ul>
      <% 
        End If
      
        If Len(vMods_SkillSet) > 0 Then 
            aSkillSet = Split (vMods_SkillSet,"::")
      %>
      <h2>
      <!--webbot bot='PurpleText' PREVIEW='Skill Set'--><%=fPhra(000241)%></h2>
      <ul>
        <% For i = 0 to Ubound(aSkillSet) %>
        <li><%=aSkillSet(i)%></li> 
        <% Next %>
      </ul>
        <% End If %> 
      <p>&nbsp;</td>
    </tr>
    </table>

  <p align="center"> <a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a></p>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->


  </body>
</html>




