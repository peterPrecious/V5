<%@ codepage=65001 %>

<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<html>

<head>
  <title>RTE_ModsDecs</title>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <script src="/V5/Inc/jQuery.js"></script>
</head>

<body>

  <%	
    Dim vModsId, aGoals, aSkillSet
    vModsId = Request("vModsId")
    sGetMods vModsId
  %> 

  <table class="table">
    <tr>
      <td>
      <h1><br><%=vMods_Title %></h1>
      <h2><!--[[-->Module Id<!--]]--> : <%=vMods_Id %></h2>

      <%=vMods_Desc%> 
          
      <% If Len(vMods_Outline)>0 Then %>
      <h3><!--[[-->Module Outline<!--]]--></h3>
      <p id="pooh"><%=vMods_Outline%></p>
      <% 
  			End If 	
  
        If Len(vMods_Goals) > 0 Then 
          aGoals = Split (vMods_Goals,"::")
      %>
      <h2><!--[[-->Learning Objectives<!--]]--></h2>
      <!--[[-->On completion of this<!--]]-->&nbsp;<%=vMods_Length%>&nbsp;<!--[[-->hour module you should be able to:<!--]]--> <br>
      <br>
      <ul>
        <% For i = 0 to Ubound(aGoals) %>
        <li><%=aGoals(i)%></li> <% Next %> </li>
      </ul>
      <% 
  			End If 
  
        If Len(vMods_SkillSet) > 0 Then 
          aSkillSet = Split (vMods_SkillSet,"::")
      %>
      <h2>
      <!--[[-->Skill Set<!--]]--></h2>
      <ul>
        <% For i = 0 to Ubound(aSkillSet) %>
        <li><%=aSkillSet(i)%></li> <% Next %> </li>
      </ul>
      <% 
  			End If 
  
        If vMods_FeaAcc Or vMods_FeaAud Or vMods_FeaMob Or vMods_FeaVid Then 
      %>
      <h2>
      <!--[[-->Features<!--]]--></h2>
      <ul>
        <table class="table">
          <% If vMods_FeaAcc Then %>
          <tr>
            <td style="width:50px;"><a href="#" title="<!--[[-->Includes compatibility with most screen readers and closed captioning (WCAG Level AA).<!--]]-->"><img border="0" src="../Images/RTE/ModsFeaAcc.png" width="16" height="16"></a></td>
            <td><!--[[-->Includes compatibility with most screen readers and closed captioning (WCAG Level AA).<!--]]--></td>
          </tr>
          <% End If %> <% If vMods_FeaAud Then %>
          <tr>
            <td style="width:50px;"><a href="#" title="<!--[[-->Requires headphones or speaker to hear audio.<!--]]-->"><img border="0" src="../Images/RTE/ModsFeaAud.png" width="16" height="16"></a></td>
            <td><!--[[-->Requires headphones or speaker to hear audio.<!--]]--></td>
          </tr>
          <% End If %> <% If vMods_FeaMob Then %>
          <tr>
            <td style="width:50px;"><a href="#" title="<!--[[-->iPad compatible, content displayed in HTML5, does not contain Flash.<!--]]-->"><img border="0" src="../Images/RTE/ModsFeaMob.png" width="16" height="16"></a></td>
            <td><!--[[-->iPad compatible, content displayed in HTML5, does not contain Flash.<!--]]--></td>
          </tr>
          <% End If %> <% If vMods_FeaVid Then %>
          <tr>
            <td style="width:50px;"><a href="#" title="<!--[[-->Contains or streams video content.<!--]]-->"><img border="0" src="../Images/RTE/ModsFeaVid.png" width="16" height="16"></a></td>
            <td><!--[[-->Contains or streams video content.<!--]]--></td>
          </tr>
          <% End If %>
        </table>
      </ul>
      <% 
  			End If 
      %> 

      </td>
    </tr>
  </table>

</body>

</html>
