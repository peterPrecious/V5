<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
    Server.Execute vShellHi
  
    Dim vStrDate, vEndDate, vStrDateErr, vEndDateErr

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
      Response.Redirect "LogReport2_X.asp?vStrDate=" & Server.UrlEncode(vStrDate) & "&vEndDate=" & Server.UrlEncode(vEndDate)
    End If

    If Request.Form("vHidden").Count = 0 Or vStrDateErr <> "" Or vEndDateErr <> "" Then
  
  %>
  <form method="POST" action="LogReport2.asp">
    <input type="hidden" name="vHidden" value="Hidden">
    <table border="1" width="100%" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#DDEEF9">
      <tr>
        <td colspan="2" align="center">
        <h1 align="center">Module Access Report - Details</h1>
        <h2>This report displays a list of learners and the modules they have last accessed during the selected period by module number.</h2>
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
  <table border="1" width="100%" cellpadding="3" cellspacing="0" style="border-collapse: collapse" bordercolor="#DDEEF9">
    <tr>
      <td colspan="5" align="center">
      <h1>
      <!--webbot bot='PurpleText' PREVIEW='Module Access Report - Details'--><%=fPhra(000175)%></h1>
      <h2 align="left">This report displays a list of learners and the modules that have been last accessed <% If Len(vStrDate) = 0 And Len(vEndDate) = 0 Then %> between January 1st of last year and <%=fFormatSqlDate(Now)%>. <% ElseIf Len(vStrDate) = 0 And Len(vEndDate) > 0 Then %> between January 1st of last year and <%=vEndDate%>. <% ElseIf Len(vStrDate) > 0 And Len(vEndDate) = 0 Then %> between <%=vStrDate%> and <%=fFormatSqlDate(Now)%>. <% Else %> between <%=vStrDate%> and <%=vEndDate%>. <% End If %><br>Note: <b>Group</b> will be generally be empty unless <b>My Learning</b> is configured.</h2>
      <p align="left">&nbsp;</p></td>
    </tr>
    <tr>
      <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap align="left">Group</th>
      <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap align="left">Learner</th>
      <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap align="left"><%=fIf(svCustPwd, fPhraH(000411), fPhraH(000211))%></th>
      <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap align="left">Module</th>
      <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" nowrap align="left">Title</th>
    </tr>
    <%
      '...get log info
      Dim vId, vIdPrev, vCriteria, vModule, vTitle, vName, vOk, vLevel
      vIDprev = ""
  
      vSql = "SELECT Memb.Memb_LastName + ',  ' + Memb.Memb_FirstName AS Name, Memb.Memb_Id AS [Id], Memb.Memb_Criteria AS [Criteria], SUBSTRING(Logs.Logs_Item, 9, 6) AS MODULE, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_Level "
      vSql = vSQL & " FROM Memb WITH (nolock) INNER JOIN Logs WITH (nolock) ON Memb.Memb_No = Logs.Logs_MembNo"
      vSql = vSQL & " WHERE (Logs.Logs_AcctID = '" & svCustAcctId & "') AND (Logs.Logs_Type = 'P') AND (Memb.Memb_Level < 4)"
      If Len(vStrDate) > 0 Then    
        vSql = vSql & " AND (Logs_Posted >= '" & vStrDate & "')"
      End If
      If Len(vEndDate) > 0 Then    
        vSql = vSql & " AND (Logs_Posted <= '" & vEndDate & "')"
      End If
      vSql = vSQL & " GROUP BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_LastName + ', ' + Memb.Memb_FirstName, Memb.Memb_Id,  Memb_Level, SUBSTRING(Logs.Logs_Item, 9, 6) "
      vSql = vSQL & " ORDER BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_ID "
  
  '   sDebug
      sOpenDb
      Set oRs = oDb.Execute(vSql)
  
      Do While Not oRS.eof

        '...ensure you can only see members with same criteria
        If svMembLevel > 2 And svMembCriteria <> "0" And oRs("Criteria") <> svMembCriteria Then 
          vOk = False
        Else
          vOk = True
        End If

        If vOk Then
  
          vId       = Trim(oRs("Id"))
          vLevel    = oRs("Memb_Level")
          vCriteria = Trim(oRs("Criteria"))
          vModule   = oRs("Module")
          vTitle    = fModsTitle (vModule)
          vName     = Trim(oRs("Name"))
          If vName  = "," Then vName = ""
    
          '...put a space between different users
          If vId  <> vIdPrev Then 
    %>
    <tr>
      <td colspan="2">&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
      <td>&nbsp;</td>
    </tr>
    <tr>
      <td><%=fif(vCriteria="0", "", fCriteria(vCriteria))%> </td>
      <td><%=fLeft(vName, 32)%> </td>
      <td><%=fId(vId, vLevel)%> </td>
      <td><%=vModule%> </td>
      <td><%=vTitle%></td>
    </tr>
    <%
		     Else 
    %>
    <tr>
      <td colspan="2">&nbsp;</td>
      <td>&nbsp;</td>
      <td><%=vModule%> </td>
      <td><%=vTitle%></td>
    </tr>
    <%
          End If
    
          vIdPrev = vId
        
        End If
        oRs.MoveNext	        
      Loop
      sCloseDB
    %>
    <tr>
      <td colspan="5" align="center"><br><a href="LogReport2.asp?vStrDate=<%=vStrDate%>&vEndDate=<%=vEndDate%>"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><br>&nbsp;</td>
    </tr>
  </table>
  <%
    End If
  
    Server.Execute vShellLo 

    Function fId (vId, vLevel)
      fId = fIf(vLevel > 2, "******", fDefault(vId, "N/A"))
    End Function

  %>

</body>

</html>


