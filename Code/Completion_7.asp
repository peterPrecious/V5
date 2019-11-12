<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include file = "Completion_Routines.asp"-->

<%
  Dim vOrder
  vOrder = fDefault(Request("vOrder"), "Id")
  Session("Completion_vProgId") = fDefault(Request("vProgId"), Session("Completion_vProgId")) 
  Session("Completion_vProgTitle") = fDefault(Request("vProgTitle"), Session("Completion_vProgTitle"))
%>

<html>

<head>
  <title>Completion_7</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>
  <% Server.Execute vShellHi %>

  <div>
    <h1><!--webbot bot='PurpleText' PREVIEW='Completion Report'--><%=fPhra(000863)%></h1>
    <h2><!--webbot bot='PurpleText' PREVIEW='Module Completion Status of<br>Program'--><%=fPhra(001631)%>&nbsp;<%=Session("Completion_vProgId") & " : " & Session("Completion_vProgTitle")%> for<br><%=Session("Completion_Learner") & " - " & Session("Completion_Name")%></h2>
  </div>


  <table class="table">
    <tr>
      <th style="width: 50%"><%=fPhraId(Session("Completion_L1tit"))%> :</th>
      <td style="width: 50%" class="c3"> <%=Session("Completion_L1val") & " : " & fL1Title(Session("Completion_L1val"))%></td>
    </tr>
    <tr>
      <th style="width: 50%"><%=fPhraId(Session("Completion_L0tit"))%> :</th>
      <td style="width: 50%" class="c3"> <%=Session("Completion_L0val") & " : " & fL0Title(Session("Completion_L0val"))%></td>
    </tr>
    <tr>
      <th ><!--webbot bot='PurpleText' PREVIEW='Roles'--><%=fPhra(000615)%> :</th>
      <td class="c3"">
        <% 
          If Len(Session("Completion_RoleD")) < 500 Then                
            Response.Write Session("Completion_RoleD")
          Else
        %>
        <a class="c3" onclick="toggle('divRoles')" href="#"><!--webbot bot='PurpleText' PREVIEW='Click here'--><%=fPhra(000876)%></a><div class="div" id="divRoles"><table class="table"><tr><td><%=Session("Completion_RoleD")%></td></tr></table></div>
        <% 
          End If 
        %>              
      </td>
    </tr>
    <tr>
      <th><!--webbot bot='PurpleText' PREVIEW='Programs | Modules'--><%=fPhra(001238)%> :</th>
      <td class="c3">
        <% 
          If Len(Session("Completion_ProgramD")) < 500 Then             
            Response.Write Session("Completion_ProgramD")
          Else
        %>
        <a class="c3" onclick="toggle('divModules')" href="#"><!--webbot bot='PurpleText' PREVIEW='Click here'--><%=fPhra(000876)%></a><div class="div" id="divModules"><table class="table"><tr><td><%=Session("Completion_ProgramD")%></td></tr></table></div>
        <% 
          End If 
        %>              
      </td>
    </tr>
  </table>

  <table style="width:600px; margin:20px auto 20px auto">
    <tr>
      <td valign="top" colspan="7" align="center">
      </td>
    </tr>
    <tr>
      <th class="rowshade">&nbsp;</th>
      <th class="rowshade"><a href="Completion_7.asp?vOrder=Id&vProgId=<%=Session("Completion_vProgId")%>&vProgTitle=<%=Session("Completion_vProgTitle")%>"><!--webbot bot='PurpleText' PREVIEW='Module'--><%=fPhra(000272)%></a></th>
      <th class="rowshade"><a href="Completion_7.asp?vOrder=Title&vProgId=<%=Session("Completion_vProgId")%>&vProgTitle=<%=Session("Completion_vProgTitle")%>"><!--webbot bot='PurpleText' PREVIEW='Title'--><%=fPhra(000019)%></a></th>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='No<br>Attempts'--><%=fPhra(000624)%></th>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='Best<br>Score'--><%=fPhra(000608)%> %</th>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='Last<br>Attempt'--><%=fPhra(000609)%></th>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='Complete?'--><%=fPhra(000606)%></th>
    </tr>
    <%
      Dim vModsTitle, vCnt
      vCnt = 0
      '...display all modules
      vSql = " SELECT "_
            & "   vRept.RepC_ModsId, "_ 
            & "   Mods.Mods_Title, "_ 
            & "   vRept.RepS_NoAttempts, "_ 
            & "   vRept.RepS_BestScore, "_ 
            & "   vRept.RepS_BestDate, "_ 
            & "   vRept.RepS_Completed "_
            & " FROM "_         
            & "   V5_Comp.dbo.vRept AS vRept WITH (NOLOCK) INNER JOIN "_
            & "   V5_Base.dbo.Mods  AS Mods  WITH (NOLOCK) ON vRept.RepC_ModsId + '" & svLang & "' = Mods.Mods_Id "_
            & " WHERE "_ 
            & "   (vRept.RepL_UserNo = " & svMembNo & ") AND "_
            & "   (vRept.RepL_MembId = '" & Session("Completion_Learner") & "') AND " _       
            & "   (vRept.RepC_ProgId = '" & Session("Completion_vProgId") & "') "_
            & " ORDER BY "_
            &     fIf(vOrder = "Id", "vRept.RepC_ModsId", "Mods.Mods_Title")
      sCompletion_Debug
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      Do While Not oRs.Eof
        vCnt = vCnt + 1
    %>
    <tr>
      <td><%=vCnt%></td>
      <td style="white-space:nowrap; text-align:center"><%=oRs("RepC_ModsId")%></td>
      <td style="white-space:nowrap;"><%=oRs("Mods_Title")%></td>
      <td style="white-space:nowrap; text-align:center"><%=oRs("RepS_NoAttempts")%></td>
      <td style="white-space:nowrap; text-align:center"><%=oRs("RepS_BestScore")%></td>
      <td style="white-space:nowrap; text-align:center"><%=fFormatDate(oRs("RepS_BestDate"))%></td>
      <td style="white-space:nowrap; text-align:center"><img border="0" src="../Images/Icons/<%=fIf(oRs("RepS_Completed"), "Check", "X")%>mark.gif" width="12" height="12"></td>
    </tr>
    <%    
        oRs.MoveNext
      Loop
      Set oRs = Nothing
      sCloseDb
    %>  

    <tr>
      <td colspan="7">&nbsp;</td>
    </tr>


    <tr>
      <td colspan="7" style="text-align:center;">
      <br><br>
      <%
        '...this generates raw data in excel
        Dim vTit, vHdr, vUrl
        vTit = "Completion Report - Program " & Session("Completion_vProgId") & " : " & Session("Completion_vProgTitle")
        vHdr = "Module ID|Title|#Attempts|Best Score %|Best Date|Complete?"
        vTit = Server.UrlEncode(vTit)
        vHdr = Server.UrlEncode(vHdr)
        vSql = Server.UrlEncode(vSql)
        vUrl = "Excel.asp?vTit=" & vTit & "&vHdr=" & vHdr & "&vSql=" & vSql
      %>
      <form name="fForm">
        <input type="button" onclick="location.href='Completion_4.asp'" value="<%=bReturn%>" name="bReturn" id="bReturn"class="button100"> 
        <input type="button" onclick="location.href='Completion_0.asp'" value="<%=bRestart%>" name="bRestart" class="button100"> 
        <input type="button" onclick="jPrint();" value="<%=bPrint%>" name="bPrint" id="bPrint" class="button100">
        <input type="button" onclick="location.href='<%=vUrl%>';" value="Excel" name="bExcel" id="bExcel" class="button100"><p><!--webbot bot='PurpleText' PREVIEW='Excel Output contains the raw data used above.'--><%=fPhra(001378)%></p>
      </form>
      </td>
    </tr>
  </table>
  

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>

