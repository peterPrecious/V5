<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include file = "Completion_Routines.asp"-->

<%
  Dim vModsId, vModsTitle, vMembId, vMembFirstName, vMembLastName, vNoNotCompleted, vOk, vOrder
  Dim aProgs, vProgs, vActive, vRoles, vAll, vYes, vCnt

  vOrder = fDefault(Request("vOrder"), "Id")
  Session("Completion_L0val") = fDefault(Request("vL0"), Session("Completion_L0val")) 
  
  p1 = Session("Completion_L0val")
  p2 = fL0Title (Session("Completion_L0val"))
%>

<html>

<head>
  <title>Completion_5</title>
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
    <h2><!--webbot bot='PurpleText' PREVIEW='Learning Completion Rates for'--><%=fPhra(001249)%><br><!--webbot bot='PurpleText' PREVIEW='^1 : ^2'--><%=fPhra(001633)%></h2>
  </div>

  <table class="table">
    <tr>
      <th style="width:50%"><%=Session("Completion_L1tit")%> :</th>
      <td class="c3" style="width:50%"><%=Session("Completion_L1val") & " : " & fL1Title(Session("Completion_L1val"))%></td>
    </tr>
    <tr>
      <th style="width:50%"><!--webbot bot='PurpleText' PREVIEW='Roles'--><%=fPhra(000615)%> :</th>
      <td class="c3" style="width:50%">
        <% 
          If Len(Session("Completion_RoleD")) < 500 Then                
            Response.Write Session("Completion_RoleD")
          Else
        %>
        <a onclick="toggle('divRoles')" href="#"><!--webbot bot='PurpleText' PREVIEW='Click here'--><%=fPhra(000876)%></a><div class="div" id="divRoles"><table class="table"><tr><td><%=Session("Completion_RoleD")%></td></tr></table></div>
        <% 
          End If 
        %>              
            
      </td>
    </tr>
    <tr>
      <th style="width:50%"><!--webbot bot='PurpleText' PREVIEW='Programs | Modules'--><%=fPhra(001238)%> :</th>
      <td class="c3"  style="width:50%">
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
      <th  style="width:50%"><!--webbot bot='PurpleText' PREVIEW='Completed'--><%=fPhra(000107)%> :</th>
      <td class="c3" style="width:50%"><%=Session("Completion_CompletedD")%></td>
    </tr>
    <tr>
      <th  style="width:50%">&nbsp;</th>
      <td class="c3"  style="width:50%">&nbsp;</td>
    </tr>
  </table>


  <table style="width:600px; margin:20px auto 20px auto">

    <tr>
      <td class="rowshade" style="width:100px">&nbsp;</td>
      <td class="rowshade" style="width:100px; text-align:left;"><a href="Completion_5.asp?vOrder=Id"><%=fPhraId(Session("Completion_LearnerId"))%></a></td>
      <td class="rowshade" style="width:200px; text-align:left;"><a href="Completion_5.asp?vOrder=Name"><!--webbot bot='PurpleText' PREVIEW='Name'--><%=fPhra(000187)%></a></td>
      <td class="rowshade" style="width:100px; text-align:center;"><!--webbot bot='PurpleText' PREVIEW='Role'--><%=fPhra(000648)%> </td>
      <td class="rowshade" style="width:100px; text-align:center;"><!--webbot bot='PurpleText' PREVIEW='Complete %'--><%=fPhra(000613)%></td>
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

'       Response.Write Session("Completion_Programs") & "<br>"
'       Response.Write vProgs & "<br>"

        '...active?
        Select Case Session("Completion_Active")
          Case "Y"   : vActive = " AND (V5_Vubz.dbo.Memb.Memb_Active = 1) "
          Case "N"   : vActive = " AND (V5_Vubz.dbo.Memb.Memb_Active = 0) "
          Case Else  : vActive = ""
        End Select

        vRoles = " AND (CHARINDEX(RIGHT(Crit.Crit_Id, 2), '" & Session("Completion_RoleP") & "') > 0)"

        '...Display all learners
        vSql = " SELECT "_     
             & "   vRept.RepL_MembId,  "_ 
             & "   vRept.RepL_MembFirstName,  "_ 
             & "   vRept.RepL_MembLastName,  "_ 
             & "   vRept.RepL_RL,  "_ 
             & "   COUNT(vRept.RepS_Completed) AS Completed_All, "_ 
             & "   SUM(CASE WHEN vRept.RepS_Completed = 1 THEN 1 ELSE 0 END) AS Completed_Yes,  "_
             & "   SUM(CASE WHEN vRept.RepS_Completed = 0 THEN 1 ELSE 0 END) AS Completed_No, "_
             & "   CAST (CAST((SUM(CASE WHEN vRept.RepS_Completed = 1 THEN 1 ELSE 0 END) * 100)  AS FLOAT(2)) / COUNT(vRept.RepS_Completed) AS FLOAT(2)) AS Percent_Yes, "_
             & "   CAST (CAST((SUM(CASE WHEN vRept.RepS_Completed = 0 THEN 1 ELSE 0 END) * 100)  AS FLOAT(2)) / COUNT(vRept.RepS_Completed) AS FLOAT(2)) AS Percent_No "_
             & " FROM "_         
             & "   V5_Comp.dbo.vRept AS vRept WITH (NOLOCK) "_
             & " WHERE "_
             & "   (vRept.RepL_UserNo = " & svMembNo & ") AND "_
             & "   (vRept.RepL_L1 = '" & Session("Completion_L1val") & "') AND "_
             & "   (vRept.RepL_L0 = '" & Session("Completion_L0val") & "') "_
             & " GROUP BY "_
             & "   vRept.RepL_MembId, vRept.RepL_MembFirstName, vRept.RepL_MembLastName, vRept.RepL_RL " _
             & " ORDER BY "_
             &    fIf(vOrder = "Id", "vRept.RepL_MembId", "vRept.RepL_MembLastName, vRept.RepL_MembFirstName")


'            & "   ((vRept.RepS_BestDate BETWEEN '" & DateAdd("d", -1, Session("Completion_StrDate")) & "' AND '" & DateAdd("d", 1, Session("Completion_EndDate")) & "') OR RepS_BestDate IS NULL) "_


        sCompletion_Debug 

        sOpenDb
        Set oRs = oDb.Execute(vSql)
        Do While Not oRs.Eof

          vMembId         = oRs("RepL_MembId")
          vMembFirstName  = oRs("RepL_MembFirstName")
          vMembLastName   = oRs("RepL_MembLastName")

          '...ok to display?
          vOk = False
          If Session("Completion_Completed") = "Y" And oRs("Completed_No") = 0 Then
            vOk = True
          ElseIf Session("Completion_Completed") = "N" And oRs("Completed_No") > 0 Then
            vOk = True
          ElseIf Session("Completion_Completed") = "X" Then
            vOk = True
          End If

          If vOk Then
            vAll = vAll + oRs("Completed_All")
            vYes = vYes + oRs("Completed_Yes")
            vCnt = vCnt + 1
      %> 
      <tr>
        <td style="white-space:nowrap; text-align:left"><%=vCnt%></td>
        <td style="white-space:nowrap; text-align:left"><a href="Completion_6.asp?vLearner=<%=vMembId%>&vName=<%= vMembFirstName & " " & vMembLastName%>"><%=vMembId%></a></td>
        <td style="white-space:nowrap; text-align:left"><%= vMembFirstName & " " & vMembLastName%></td>
        <td style="white-space:nowrap; text-align:center"><%=oRs("RepL_RL")%></td>
        <td style="white-space:nowrap; text-align:center"><%=FormatNumber(oRs("Completed_Yes")/oRs("Completed_All")*100, 0)%>%</td>
      </tr>
      <%    
          End If
          oRs.MoveNext
        Loop
        Set oRs = Nothing
        sCloseDb
        If vAll = 0 Then vYes = 0 : vAll = 1
        p1 = Session("Completion_L0tit")
      %> 

      <tr>
        <td colspan="5">&nbsp;</td>
      </tr>
      <tr>
        <th colspan="4"><!--webbot bot='PurpleText' PREVIEW='^1 Total Completed'--><%=fPhra(001380)%> :</th>
        <th style="text-align:center;"><%=FormatNumber(vYes/vAll*100, 0)%>%</th>
      </tr>
      <tr>
      <td colspan="5" style="text-align:center;">

        <br><br>
        <% p1 = Session("Completion_L0tit") %>
        <!--webbot bot='PurpleText' PREVIEW='^1 Total Completed is the percentage of selected assessments completed.'--><%=fPhra(001381)%><br>

        <%
          '...this generates raw data in excel
          Dim vTit, vHdr, vUrl
          vTit = "Completion Report - " & Session("Completion_L0tit")
          vHdr = Session("Completion_LearnerId") & "|First Name|Last Name|Role|#Learners|#Completed|#Not Completed|%Completed|%Not Completed"
          vTit = Server.UrlEncode(vTit)
          vHdr = Server.UrlEncode(vHdr)
          vSql = Server.UrlEncode(vSql)
          vUrl = "Excel.asp?vTit=" & vTit & "&vHdr=" & vHdr & "&vSql=" & vSql
        %>

        <form name="fForm">
          <input type="button" onclick="location.href='Completion_<%=fIf(svMembLevel > 4, "4", "0")%>.asp'" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button100"> 
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

