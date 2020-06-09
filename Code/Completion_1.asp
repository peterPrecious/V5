<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include file = "Completion_Routines.asp"-->

<%
  Dim vOrder, vAll, vYes, vCnt, vTotal, aProgs, vProgs, vActive
  vOrder = fDefault(Request("vOrder"), "Id")  
%>

<html>

<head>
  <title>Completion_1</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <div>
    <h1><!--webbot bot='PurpleText' PREVIEW='Completion Report'--><%=fPhra(000863)%></h1>
  </div>

  <table style="width:600px; margin:20px auto 20px auto">
    <tr>
      <th style="width:50%"><!--webbot bot='PurpleText' PREVIEW='Learning Completion Rates'--><%=fPhra(000603)%> :</th>
      <td class="c3"  style="width:50%"><!--webbot bot='PurpleText' PREVIEW='National'--><%=fPhra(000604)%></td>
    </tr>
    <tr>
      <th style="width:50%"><!--webbot bot='PurpleText' PREVIEW='Roles'--><%=fPhra(000615)%> :</th>
      <td class="c3" style="width:50%">
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
      <th style="width:50%"><!--webbot bot='PurpleText' PREVIEW='Programs | Modules'--><%=fPhra(001238)%> :</th>
      <td class="c3" style="width:50%">
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
    <tr>
      <th style="width:50%">
      <!--webbot bot='PurpleText' PREVIEW='Completed'--><%=fPhra(000107)%> :</th>
      <td class="c3" style="width:50%"><%=Session("Completion_CompletedD")%></td>
    </tr>
    <tr>
      <th style="width:50%">
      &nbsp;</th>
      <td class="c3" style="width:50%">&nbsp;</td>
    </tr>
  </table>


  <table style="width:600px; margin:20px auto 20px auto">

    <tr>
      <td class="rowshade" style="width:100px">&nbsp;</td>
      <td class="rowshade" style="width:300px; text-align:left;"><a href="Completion_1.asp?vOrder=Id"><%=fPhraId(Session("Completion_L1tit"))%></a></td>
      <td class="rowshade" style="width:100px; text-align:left;"><a href="Completion_1.asp?vOrder=Name"><!--webbot bot='PurpleText' PREVIEW='Title'--><%=fPhra(000019)%></a></td>
      <td class="rowshade" style="width:090px; text-align:center;"><!--webbot bot='PurpleText' PREVIEW='Complete %'--><%=fPhra(000613)%></td>
    </tr>
    <%
      vAll = 0 : vYes = 0 : vCnt = 0 : vTotal = 0

      '...get programs (P1234|1234|2323|1212,P1235|1239) - strip off modules
      aProgs = Split(Session("Completion_Programs"), ",")
      vProgs = ""
      For i = 0 To Ubound(aProgs)
          vProgs = vProgs& Left(aProgs(i), 5) & "," 
      Next
      vProgs = Left(vProgs, Len(vProgs)-1) 
      vProgs = " AND (Comp_ProgId IN ('" & Replace(vProgs, ",", "','") & "')) "

      '...Count Learners in all L1s
      vSql = " SELECT "_
            & "   COUNT(*) AS Total "_
            & " FROM "_         
            & "   V5_Comp.dbo.RepL WITH (NOLOCK) "_
            & " WHERE "_
            & "   RepL_UserNo = " & svMembNo

      sCompletion_Debug
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      vTotal = oRs("Total")
      Set oRs = Nothing

      vSql = " SELECT "_     
            & "   vRept.RepL_L1, "_ 
            & "   vRept.Unit_L1Title, "_ 
            & "   COUNT(vRept.RepS_Completed) AS Completed_All, "_ 
            & "   SUM(CASE WHEN vRept.RepS_Completed = 1 THEN 1 ELSE 0 END) AS Completed_Yes, "_ 
            & "   SUM(CASE WHEN vRept.RepS_Completed = 0 THEN 1 ELSE 0 END) AS Completed_No, "_ 
            & "   CAST(CAST(SUM(CASE WHEN vRept.RepS_Completed = 1 THEN 1 ELSE 0 END) * 100 AS FLOAT(2)) / COUNT(vRept.RepS_Completed) AS FLOAT(2)) AS Percent_Yes, "_ 
            & "   CAST(CAST(SUM(CASE WHEN vRept.RepS_Completed = 0 THEN 1 ELSE 0 END) * 100 AS FLOAT(2)) / COUNT(vRept.RepS_Completed) AS FLOAT(2)) AS Percent_No "_
            & " FROM "_         
            & "   V5_Comp.dbo.vRept AS vRept WITH (NOLOCK) "_
            & " WHERE "_
            & "   (vRept.RepL_UserNo = " & svMembNo & ") "_
            & " GROUP BY "_ 
            & "   vRept.RepL_L1, vRept.Unit_L1Title "_
            & " ORDER BY "_ 
            &     fIf(vOrder = "Id", "vRept.RepL_L1", "vRept.Unit_L1Title")


'            & "   ((vRept.RepS_BestDate BETWEEN '" & DateAdd("d", -1, Session("Completion_StrDate")) & "' AND '" & DateAdd("d", 1, Session("Completion_EndDate")) & "') OR RepS_BestDate IS NULL) "_


      sCompletion_Debug

      sOpenDb
      Set oRs = oDb.Execute(vSql)
      Do While Not oRs.Eof
        vAll = vAll + oRs("Completed_All")
        vYes = vYes + oRs("Completed_Yes")
        vCnt = vCnt + 1
    %>
    <tr>
      <td style="white-space:nowrap; text-align:left"><%=vCnt%></td>
      <td style="white-space:nowrap; text-align:left"><a href="Completion_4.asp?vL1=<%=oRs("RepL_L1")%>"><%=oRs("RepL_L1")%></a></td>
      <td style="white-space:nowrap; text-align:left"><%=oRs("Unit_L1Title")%></td>
      <td style="white-space:nowrap; text-align:center"><%=FormatNumber(oRs("Completed_Yes")/oRs("Completed_All")*100, 0)%>%</td>
    </tr>
    <%    
        oRs.MoveNext
      Loop
      Set oRs = Nothing
      sCloseDb
      If vAll = 0 Then vYes = 0 : vAll = 100
    %>

    <tr>
      <td colspan="4">&nbsp;</td>
    </tr>
    <tr>
      <th colspan="3"><%=Session("Completion_L9tit")%> <!--webbot bot='PurpleText' PREVIEW='Total Completed'--><%=fPhra(001239)%> :</th>
      <th style="text-align:center;"><%=FormatNumber(vYes/vAll*100, 0)%>%</th>
    </tr>
    <tr>
      <th colspan="3"><!--webbot bot='PurpleText' PREVIEW='Number of Learners Included'--><%=fPhra(000647)%>&nbsp;:</th>
      <th style="text-align:center;"><%=vTotal%>&nbsp; </th>
    </tr>
    <tr>
      <td colspan="4" style="text-align:center;">

      <br><br><%=Session("Completion_L9tit")%>&nbsp;<!--webbot bot='PurpleText' PREVIEW='Total Completed is the percentage of selected assessments completed.'--><%=fPhra(001240)%><br>
      <%
        '...this generates raw data in excel
        Dim vTit, vHdr, vUrl
        vTit = "Completion Report - National"
        vHdr = "Prov/State|Title|#Learners|#Completed|#Not Completed|%Completed|%Not Completed"
        vTit = Server.UrlEncode(vTit)
        vHdr = Server.UrlEncode(vHdr)
        vSql = Server.UrlEncode(vSql)
        vUrl = "Excel.asp?vTit=" & vTit & "&vHdr=" & vHdr & "&vSql=" & vSql
      %>
      <form name="fForm">
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

