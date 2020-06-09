<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include file = "Completion_Routines.asp"-->

<%
  Dim vOrder, vCnt, vAll, vYes, aProgs, vProgs
  vOrder = fDefault(Request("vOrder"), "Id")
  Session("Completion_Learner") = fDefault(Request("vLearner"), Session("Completion_Learner"))
  Session("Completion_Name")    = fDefault(Request("vName"), Session("Completion_Name"))
%>

<html>

<head>
  <title>Completion_6</title>
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
    <h2>
      <!--webbot bot='PurpleText' PREVIEW='Program Completion Rates'--><%=fPhra(000775)%> for: <br>
      <%=Session("Completion_Learner") & " - " & Session("Completion_Name")%>
    </h2>
  </div>

  <table class="table">
     <tr>
      <th style="width:50%"><%=fPhraId(Session("Completion_L1tit"))%> :</th>
      <td style="width:50%" class="c3"><%=Session("Completion_L1val") & " : " & fL1Title(Session("Completion_L1val"))%></td>
    </tr>
    <tr>
      <th style="width:50%"><%=fPhraId(Session("Completion_L0tit"))%> :</th>
      <td style="width:50%" class="c3"><%=Session("Completion_L0val") & " : " & fL0Title(Session("Completion_L0val"))%></td>
    </tr>
    <tr>
      <th style="width:50%"><!--webbot bot='PurpleText' PREVIEW='Completed'--><%=fPhra(000107)%> :</th>
      <td style="width:50%" class="c3"><%=Session("Completion_CompletedD")%></td>
    </tr>
    <tr>
      <th><!--webbot bot='PurpleText' PREVIEW='Roles'--><%=fPhra(000615)%> :</th>
      <td class="c3">
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
      <td class="rowshade" style="width:100px">&nbsp;</td>
      <th class="rowshade" style="width:150px; text-align:center;"><a href="Completion_6.asp?vOrder=Id&vLearner=<%=Session("Completion_Learner")%>&vName=<%=Session("Completion_Name")%>"><!--webbot bot='PurpleText' PREVIEW='Program'--><%=fPhra(000201)%></a></th>
      <th class="rowshade" style="width:250px; text-align:left;"><a href="Completion_6.asp?vOrder=Title&vLearner=<%=Session("Completion_Learner")%>&vName=<%=Session("Completion_Name")%>"><!--webbot bot='PurpleText' PREVIEW='Title'--><%=fPhra(000019)%></a></th>
      <th class="rowshade" style="width:090px; text-align:center;"><!--webbot bot='PurpleText' PREVIEW='Complete %'--><%=fPhra(000613)%></th>
    </tr>
    <%
      vAll = 0 : vYes = 0 : vCnt = 0

      '...get programs (P1234|1234|2323|1212,P1235|1239) - strip off modules
      aProgs = Split(Session("Completion_Programs"), ",")
      vProgs = ""
      For i = 0 To Ubound(aProgs)
          vProgs = vProgs& Left(aProgs(i), 5) & "," 
      Next
      vProgs = Left(vProgs, Len(vProgs)-1) 
      vProgs = " AND (Comp_ProgId IN ('" & Replace(vProgs, ",", "','") & "')) "

      '...display all programs
      vSql = " SELECT "_
            & "   vRept.RepL_MembId, "_ 
            & "   vRept.RepC_ProgId, "_ 
            & "   Prog.Prog_Title1 AS ProgTitle, "_
            & "   COUNT(vRept.RepS_Completed) AS Completed_All, "_ 
            & "   SUM(CASE WHEN vRept.RepS_Completed = 1 THEN 1 ELSE 0 END) AS Completed_Yes,  "_
            & "   SUM(CASE WHEN vRept.RepS_Completed = 0 THEN 1 ELSE 0 END) AS Completed_No, "_
            & "   CAST (CAST((SUM(CASE WHEN vRept.RepS_Completed = 1 THEN 1 ELSE 0 END) * 100)  AS FLOAT(2)) / COUNT(vRept.RepS_Completed) AS FLOAT(2)) AS Percent_Yes, "_
            & "   CAST (CAST((SUM(CASE WHEN vRept.RepS_Completed = 0 THEN 1 ELSE 0 END) * 100)  AS FLOAT(2)) / COUNT(vRept.RepS_Completed) AS FLOAT(2)) AS Percent_No "_
            & " FROM "_         
            & "   V5_Comp.dbo.vRept AS vRept WITH (NOLOCK) INNER JOIN "_
            & "   V5_Base.dbo.Prog  AS  Prog WITH (NOLOCK) ON vRept.RepC_ProgId + '" & svLang & "' = Prog.Prog_Id "_
            & " WHERE "_ 
            & "   (vRept.RepL_UserNo = " & svMembNo & ") AND "_
            & "   (vRept.RepL_MembId = '" & Session("Completion_Learner") & "') " _       
            & " GROUP BY " _
            & "   vRept.RepL_MembId, "_ 
            & "   vRept.RepC_ProgId, "_ 
            & "   Prog.Prog_Title1 " _
            & " ORDER BY "_
            &     fIf(vOrder = "Id", "vRept.RepC_ProgId", "Prog.Prog_Title1")
      sCompletion_Debug
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      Do While Not oRs.Eof
        vCnt = vCnt + 1

        vAll = vAll + oRs("Completed_All")
        vYes = vYes + oRs("Completed_Yes")
          
    %>
    <tr>
      <td style="white-space:nowrap; text-align:left"><%=vCnt%></td>
      <td style="white-space:nowrap; text-align:center"><a href="Completion_7.asp?vProgId=<%=oRs("RepC_ProgId")%>&vProgTitle=<%=Server.UrlEncode(oRs("ProgTitle"))%>"><%=oRs("RepC_ProgId")%></a></td>
      <td style="white-space:nowrap; text-align:left"><%= fLeft(oRs("ProgTitle"), 32)%></td>
      <td style="white-space:nowrap; text-align:center">
      <% 
        If oRs("Completed_All") = 0 Then
          Response.Write "0%"
        Else
          Response.Write FormatPercent(oRs("Completed_Yes") / oRs("Completed_All"), 0) 
        End If
      %>
      </td>
    </tr>
    <%    
        oRs.MoveNext
      Loop
      Set oRs = Nothing
      sCloseDb
    %>  

    <tr>
      <td colspan="4">&nbsp;</td>
    </tr>

    <tr>
      <th colspan="3"><!--webbot bot='PurpleText' PREVIEW='Program Total Completed'--><%=fPhra(000776)%> :</th>
      <th style="text-align:center;"><%=FormatNumber(vYes/vAll*100, 0)%>%</th>
    </tr>

    <tr>
      <td colspan="4" style="text-align:center;">
      <br><br>
      <!--webbot bot='PurpleText' PREVIEW='Program Total Completed is the percentage of selected assessments completed.'--><%=fPhra(000927)%><br>

      <%
        '...this generates raw data in excel
        Dim vTit, vHdr, vUrl
        vTit = "Completion Report - " & Session("Completion_Name")
        vHdr = Session("Completion_LearnerId") & "|Program Id|Title|#Learners|#Completed|#Not Completed|%Completed|%Not Completed"
        vTit = Server.UrlEncode(vTit)
        vHdr = Server.UrlEncode(vHdr)
        vSql = Server.UrlEncode(vSql)
        vUrl = "Excel.asp?vTit=" & vTit & "&vHdr=" & vHdr & "&vSql=" & vSql
      %>


      <form name="fForm">
        <input type="button" onclick="location.href='Completion_5.asp'" value="<%=bReturn%>" name="bReturn" id="bReturn"class="button100"> 
        <input type="button" onclick="location.href='Completion_0.asp'" value="<%=bRestart%>" name="bRestart" class="button100"> 
        <input type="button" onclick="jPrint();" value="<%=bPrint%>" name="bPrint" id="bPrint" class="button100">
        <input type="button" onclick="location.href='<%=vUrl%>';" value="Excel" name="bExcel" id="bExcel" class="button100">
        <p><!--webbot bot='PurpleText' PREVIEW='Excel Output contains the raw data used above.'--><%=fPhra(001378)%></p>
      </form>
      </td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>

