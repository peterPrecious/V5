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
      <h2><!--webbot bot='PurpleText' PREVIEW='Module Id'--><%=fPhra(001271)%> : <%=vMods_Id %></h2>

      <%=vMods_Desc%> 
          
      <% If Len(vMods_Outline)>0 Then %>
      <h3><!--webbot bot='PurpleText' PREVIEW='Module Outline'--><%=fPhra(000177)%></h3>
      <p id="pooh"><%=vMods_Outline%></p>
      <% 
  			End If 	
  
        If Len(vMods_Goals) > 0 Then 
          aGoals = Split (vMods_Goals,"::")
      %>
      <h2><!--webbot bot='PurpleText' PREVIEW='Learning Objectives'--><%=fPhra(000169)%></h2>
      <!--webbot bot='PurpleText' PREVIEW='On completion of this'--><%=fPhra(000417)%>&nbsp;<%=vMods_Length%>&nbsp;<!--webbot bot='PurpleText' PREVIEW='hour module you should be able to:'--><%=fPhra(000418)%> <br>
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
      <!--webbot bot='PurpleText' PREVIEW='Skill Set'--><%=fPhra(000241)%></h2>
      <ul>
        <% For i = 0 to Ubound(aSkillSet) %>
        <li><%=aSkillSet(i)%></li> <% Next %> </li>
      </ul>
      <% 
  			End If 
  
        If vMods_FeaAcc Or vMods_FeaAud Or vMods_FeaMob Or vMods_FeaVid Then 
      %>
      <h2>
      <!--webbot bot='PurpleText' PREVIEW='Features'--><%=fPhra(001367)%></h2>
      <ul>
        <table class="table">
          <% If vMods_FeaAcc Then %>
          <tr>
            <td style="width:50px;"><a href="#" title="<!--webbot bot='PurpleText' PREVIEW='Includes compatibility with most screen readers and closed captioning (WCAG Level AA).'--><%=fPhra(001442)%>"><img border="0" src="../Images/RTE/ModsFeaAcc.png" width="16" height="16"></a></td>
            <td><!--webbot bot='PurpleText' PREVIEW='Includes compatibility with most screen readers and closed captioning (WCAG Level AA).'--><%=fPhra(001442)%></td>
          </tr>
          <% End If %> <% If vMods_FeaAud Then %>
          <tr>
            <td style="width:50px;"><a href="#" title="<!--webbot bot='PurpleText' PREVIEW='Requires headphones or speaker to hear audio.'--><%=fPhra(001443)%>"><img border="0" src="../Images/RTE/ModsFeaAud.png" width="16" height="16"></a></td>
            <td><!--webbot bot='PurpleText' PREVIEW='Requires headphones or speaker to hear audio.'--><%=fPhra(001443)%></td>
          </tr>
          <% End If %> <% If vMods_FeaMob Then %>
          <tr>
            <td style="width:50px;"><a href="#" title="<!--webbot bot='PurpleText' PREVIEW='iPad compatible, content displayed in HTML5, does not contain Flash.'--><%=fPhra(001444)%>"><img border="0" src="../Images/RTE/ModsFeaMob.png" width="16" height="16"></a></td>
            <td><!--webbot bot='PurpleText' PREVIEW='iPad compatible, content displayed in HTML5, does not contain Flash.'--><%=fPhra(001444)%></td>
          </tr>
          <% End If %> <% If vMods_FeaVid Then %>
          <tr>
            <td style="width:50px;"><a href="#" title="<!--webbot bot='PurpleText' PREVIEW='Contains or streams video content.'--><%=fPhra(001445)%>"><img border="0" src="../Images/RTE/ModsFeaVid.png" width="16" height="16"></a></td>
            <td><!--webbot bot='PurpleText' PREVIEW='Contains or streams video content.'--><%=fPhra(001445)%></td>
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


