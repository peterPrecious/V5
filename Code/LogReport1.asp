<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <base target="_self">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
</head>

<body>

  <% 
    Server.Execute vShellHi

    Dim vStrDate, vEndDate, vStrDateErr, vEndDateErr
    Dim vId, vIdPrev, vLevel, vLevelPrev, vIdLast, vModule, vModules, vModulePrev, vCriteriaNo, vCriteriaId, vName, vNamePrev, vOk
  

    '...default to previous month
    If Request("vStrDate").Count = 0 And Request("vEndDate").Count = 0 Then
      vStrDateErr = "" : vStrDate = fFormatSqlDate(MonthName(Month(Now)) & " 1, " & Year(Now))
      vEndDateErr = "" : vEndDate = fFormatSqlDate(DateAdd("d", -1, MonthName(Month(DateAdd("m", +1, Now))) & " 1, " & Year(DateAdd("m", +1, Now))))
    Else
      vStrDate  = fFormatSqlDate(Request("vStrDate")) 
      If Request("vStrDate") = "" Then 
        vStrDate = ""
      ElseIf vStrDate = " " Then
        vStrDate  = Request("vStrDate") 
        vStrDateErr = "Error"
      End If
      vEndDate  = fFormatSqlDate(Request("vEndDate"))
      If Request("vEndDate") = "" Then 
        vEndDate = ""
      ElseIf vEndDate = " " Then
        vEndDate  = Request("vEndDate") 
        vEndDateErr = "Error"
      End If
      If (Len(vStrDate) > 0 And vStrDateErr = "") And (Len(vEndDate) > 0 And vEndDateErr = "") Then
        If DateDiff("d", vStrDate, vEndDate) < 0 Then
          vEndDateErr = "Error"
        End If
      End If
    End If

    If Request.Form("bExcel").Count = 1 And vStrDateErr = "" And vEndDateErr = "" Then
      Response.Redirect "LogReport1_X.asp?vStrDate=" & Server.UrlEncode(vStrDate) & "&vEndDate=" & Server.UrlEncode(vEndDate)
    End If
    
    If Request.Form("vHidden").Count = 0 Or vStrDateErr <> "" Or vEndDateErr <> "" Then
  %>
  <form method="POST" action="LogReport1.asp">
    <input type="Hidden" name="vHidden" value="Hidden">
    <table border="1" width="100%" cellpadding="3" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td colspan="2" align="center">
        <h1 align="center">Module Access Report - Summary</h1>
        <h2>This report displays, in learner last name order, all learners and the modules they have last accessed during the selected month by module number.</h2>
        </td>
      </tr>
      <tr>
        <th align="right" valign="top" width="35%">Select Start Date :</th>
        <td width="65%"><input type="text" name="vStrDate" size="15" value="<%=vStrDate%>"> <span style="background-color: #FFFF00"><%=vStrDateErr%></span><br>ie Jan 1, 2005 (MMM DD, YYYY). Leave empty to start at first record.&nbsp; Note: activity logs are are only maintained from January of the previous year.</td>
      </tr>
      <tr>
        <th align="right" valign="top" width="35%">End Date :</th>
        <td width="65%"><input type="text" name="vEndDate" size="15" value="<%=vEndDate%>"> <span style="background-color: #FFFF00"><%=vEndDateErr%></span><br>ie Mar 31, 2005 (MMM DD, YYYY). Leave empty to finish with last record.</td>
      </tr>
      <tr>
        <th height="100" width="100%" colspan="2">
        <input type="submit" value="Online" name="bPrint" id="bPrint" class="button070">
        or <input type="submit" value="Excel" name="bExcel" class="button070"></th>
        </tr>
    </table>
  </form>
  <%
    Else
  %>
  <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="3" cellspacing="0">
    <tr>
      <td colspan="4" align="center">
      <h1>Module Access Report - Summary</h1>
      <h2>This report displays, in learner last name order, all learners and the modules that have been last accessed<br> <% If Len(vStrDate) = 0 And Len(vEndDate) = 0 Then %> between January 1st of last year and <%=fFormatSqlDate(Now)%>. <% ElseIf Len(vStrDate) = 0 And Len(vEndDate) > 0 Then %> between January 1st of last year and <%=vEndDate%>. <% ElseIf Len(vStrDate) > 0 And Len(vEndDate) = 0 Then %> between <%=vStrDate%> and <%=fFormatSqlDate(Now)%>. <% Else %> between <%=vStrDate%> and <%=vEndDate%>. <% End If %> </h2>
      </td>
    </tr>
    <tr>
      <th nowrap align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Group</th>
      <th nowrap align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Learner</th>
      <th nowrap align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF"><%=fIf(svCustPwd, fPhraH(000411), fPhraH(000211))%></th>
      <th nowrap align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Modules accessed</th>
    </tr>
    <%
      '...get log info
      vIdPrev = "": vLevelPrev = 0: vModules = "": vIdLast = ""

      vSql = "SELECT Memb.Memb_LastName + ',  ' + Memb.Memb_FirstName AS [Name], Memb.Memb_Id AS [Id], Memb.Memb_Criteria AS [Criteria], SUBSTRING(Logs.Logs_Item, 9, 6) AS MODULE, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_Level "
      vSql = vSQL & " FROM Memb WITH (nolock) INNER JOIN Logs WITH (nolock) ON Memb.Memb_No = Logs.Logs_MembNo"
      vSql = vSQL & " WHERE (Logs.Logs_AcctId = '" & svCustAcctId & "') AND (Logs.Logs_Type = 'P') AND (Memb.Memb_Level < 5)"
      If Len(vStrDate) > 0 Then    
        vSql = vSql & " AND (Logs_Posted >= '" & vStrDate & "')"
      End If
      If Len(vEndDate) > 0 Then    
        vSql = vSql & " AND (Logs_Posted <= '" & vEndDate & "')"
      End If

      vSql = vSQL & " GROUP BY Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_Id, Memb.Memb_Criteria, Memb.Memb_LastName + Memb.Memb_FirstName, SUBSTRING(Logs.Logs_Item, 9, 6), Memb.Memb_Level "
      vSql = vSQL & " ORDER BY Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_Id  "
  
'     sDebug
      sOpenDb
      Set oRs = oDb.Execute(vSql)
  
      Do While Not oRS.Eof

        '...ensure you can only see members with same criteria
        If svMembLevel > 2 And svMembCriteria <> "0" And oRs("Criteria") <> svMembCriteria Then 
          vOk = False
        Else
          vOk = True
        End If

        If vOk Then

          vId     = Trim(oRs("Id"))
          vModule = oRs("Module")
          vLevel  = oRs("Memb_Level")
          vName   = fDefault(Trim(oRs("Name")), "N/A")

          vCriteriaNo = Trim(oRs("Criteria"))
          vCriteriaId = fIf(vCriteriaNo="0", "", fCriteria(vCriteriaNo))  
    
          '...if very first record
          If vIdPrev    = "" Then 
            vIdPrev     = vId
            vLevelPrev  = vLevel
            vModulePrev = vModule
            vModules    = vModule
            vNamePrev   = vName
          End If
           
          '...if new Id print out Id and modules
          If vIdPrev <> vId Then               
    %>
    <tr>
<!--
      <td valign="top" height="12"><%=fIf(vCriteriaNo="0", "", fCriteria(vCriteriaNo))%> </td>
-->      
      <td valign="top" height="12"><%=vCriteriaId%> </td>
      <td valign="top" height="12" nowrap><%=fLeft(vNamePrev, 24)%> </td>
      <td valign="top" height="12"><%=fId(vIdPrev, vLevelPrev)%> </td>
      <td valign="top" height="12"><%=vModules%> </td>
    </tr>
    <%
            vIdLast     = vIdPrev
            vIdPrev     = vId
            vLevelPrev  = vLevel
            vModulePrev = vModule
            vModules    = vModule 
            vNamePrev   = vName

          ElseIf vModule <> vModulePrev Then
            vModules = vModules & " " & vModule
            vModulePrev = vModule
          End If
        
        End If  
          
        oRs.MoveNext	        
      Loop
      sCloseDB
   
      '...any stragglers
      If vId <> vIdLast Then     
    %>
    <tr>
      <td valign="top" height="12"> </td>
      <td valign="top" height="12" nowrap><%=fLeft(vName, 24)%> </td>
      <td valign="top" height="12"><%=fId(vId, vLevel)%> </td>
      <td valign="top" height="12"><%=vModules%> </td>
    </tr>
    <%
	    End If  
    %>
    <tr>
      <td colspan="4" align="center" height="46"><br>&nbsp;<a href="LogReport1.asp?vStrDate=<%=vStrDate%>&vEndDate=<%=vEndDate%>"><input type="button" value="Return" name="bReturn" id="bReturn"class="button"></a><br>&nbsp; </td>
    </tr>
  </table>
  <%
    End If
    
    Function fId (vId, vLevel)
      fId = fIf(vLevel > 2, "******", fDefault(vId, "N/A"))
    End Function
    
  %>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

