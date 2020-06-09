<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
    Server.Execute vShellHi

    Dim vAccount, vLevel, vLearners, vStrDate, vEndDate, vStrDateErr, vEndDateErr
  
    vAccount  = fDefault(Request("vAccount"), "current")
    vLevel    = fDefault(Request("vLevel"), "prog")
    vLearners = fDefault(Request("vLearners"), "n")

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

    '...if Excel then go to Log0X
    If Request.Form("bExcel").Count = 1 And vStrDateErr = "" And vEndDateErr = "" Then
      Response.Redirect "LogReport0X.asp?vStrDate=" & Server.UrlEncode(vStrDate) & "&vEndDate=" & Server.UrlEncode(vEndDate) & "&vAccount=" & vAccount & "&vLevel=" & vLevel & "&vLearners=" & vLearners
    End If

    If Request.Form("vHidden").Count = 0 Or vStrDateErr <> "" Or vEndDateErr <> "" Then
  %>
  <form method="POST" action="LogReport0.asp">
    <input type="Hidden" name="vHidden" value="Hidden">
    <table border="1" width="100%" cellpadding="3" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td colspan="2" align="left">
        <h1 align="center">
        <!--[[-->Time spent in Program | Module<!--]]--></h1>
        <h2>
        <!--[[-->This report displays the total time a learner has spent in minutes in a module or program for those who have accessed content within a specified timeframe.&nbsp;&nbsp;Date parameters allow you to select a report on learners who have accessed content within a certain timeframe, for example, the month of June, so that only learners who have accessed in June will be displayed.&nbsp;&nbsp;Note that the time spent that will be displayed is the total time - not just the time spent in June - but rather the cumulative time the learner has spent overall in a particular module or program.&nbsp;&nbsp;For example, if a learner has spent 15 minutes in May and 15 minutes in June, the report for June will display a total of 30 minutes (15 in May and 15 in June), not just the 15 minutes in June.<!--]]-->&nbsp;
        <span style="background-color: #FFFF00"><!--[[-->A learning session must be greater than 1 minute to be added into this report.<!--]]--></span></h2>
        </td>
      </tr>
      <tr>
        <th align="right" valign="top" width="30%">
        <!--[[-->Select Start Date<!--]]--> :</th>
        <td width="70%"><input type="text" name="vStrDate" size="15" value="<%=vStrDate%>"> <span style="background-color: #FFFF00"><%=vStrDateErr%></span><br>
        <!--[[-->ie Jan 1, 2006 (MMM DD, YYYY). Leave empty to start at first record.&nbsp; Note: activity logs are are only maintained from Jan 1st of the previous year.<!--]]--></td>
      </tr>
      <tr>
        <th align="right" valign="top" width="30%">
        <!--[[-->End Date<!--]]--> :</th>
        <td width="70%"><input type="text" name="vEndDate" size="15" value="<%=vEndDate%>"> <span style="background-color: #FFFF00"><%=vEndDateErr%></span><br>
        <!--[[-->ie Mar 31, 2006 (MMM DD, YYYY). Leave empty to finish with last record.<!--]]--></td>
      </tr>
      <tr>
        <th align="right" nowrap width="35%" valign="top">
        <!--[[-->include<!--]]--> :</th>
        <td width="65%">
          <input type="radio" value="y" name="vLearners" <%=fcheck("y", vlearners)%>><!--[[-->Learners only<!--]]--><br>
          <input type="radio" value="n" name="vLearners" <%=fcheck("n", vlearners)%>><!--[[-->Include Facilitators<!--]]--></td>
      </tr>
      <tr>
        <th align="right" width="35%" nowrap valign="top">
        <!--[[-->Level<!--]]--> :</th>
        <td width="65%">
          <input type="radio" value="prog" name="vLevel" <%=fcheck("prog", vlevel)%>><!--[[-->Program<!--]]--><br>
          <input type="radio" value="mods" name="vLevel" <%=fcheck("mods", vlevel)%>><!--[[-->Module<!--]]--></td>
      </tr>
      <tr>
        <th align="right" height="50">and Format :</th>
        <td height="50">
          <input type="submit" value="Online" name="bOnline" class="button"> <b>or ...&nbsp; </b>
          <input type="submit" value="Excel" name="bExcel" class="button"></td>
      </tr>
      </table>
  </form>
  <%
    Else
      '...get log info
      Dim vId, vModule, vTitle, vLearner, vIDPrev, vTimeSpent, vLogItemLength, vLogItemTitle, vOk
      vIdPrev = ""

      '...if summarizing to prog level just select the left 7 chars (ie P1001EN) else select all 14 chars (ie P1001EN|1234EN)
      vLogItemLength = 14: If vLevel = "prog" Then vLogItemLength = 7
      vLogItemTitle  = "<!--{{-->Modules<!--}}-->" : If vLevel = "Prog" Then vLogItemTitle  = "<!--{{-->Programs<!--}}-->" 
  %>
  <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="2" cellspacing="0">
    <tr>
      <td colspan="5" align="left">
        <h1 align="center"><!--[[-->Time spent in Program | Module<!--]]--></h1>
        <h2><!--[[-->This report displays the total time a learner has spent in minutes in a module or program for those who have accessed content within a specified timeframe.&nbsp;&nbsp;Date parameters allow you to select a report on learners who have accessed content within a certain timeframe, for example, the month of June, so that only learners who have accessed in June will be displayed.&nbsp;&nbsp;Note that the time spent that will be displayed is the total time - not just the time spent in June - but rather the cumulative time the learner has spent overall in a particular module or program.&nbsp;&nbsp;For example, if a learner has spent 15 minutes in May and 15 minutes in June, the report for June will display a total of 30 minutes (15 in May and 15 in June), not just the 15 minutes in June.<!--]]-->&nbsp; <span style="background-color: #FFFF00"><!--[[-->A learning session must be greater than 1 minute to be added into this report.<!--]]--></span></h2>
      </td>
    </tr>
    <tr>
      <th bgcolor="#DDEEF9" align="left" bordercolor="#FFFFFF" height="30"><!--[[-->Name<!--]]--> </th>
      <th bgcolor="#DDEEF9" align="left" bordercolor="#FFFFFF" height="30"><%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%></th>
      <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">&nbsp;<!--[[-->Time<!--]]--> </th>
      <th bgcolor="#DDEEF9" align="left" bordercolor="#FFFFFF" height="30"><%=vLogItemTitle%></th>
    </tr>
    <%
      vSql = "SELECT Memb.Memb_LastName + ',  ' + Memb.Memb_FirstName AS [Learner], Memb.Memb_Id AS [Id], Memb.Memb_Criteria AS [Criteria], Left(Logs.Logs_Item, " & vLogItemLength & ") AS MODULE, SUM(CONVERT(integer, RIGHT(Logs_Item, 6))) AS TIMESPENT, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_Level "
      vSql = vSQL & " FROM Memb WITH (nolock) INNER JOIN Logs WITH (nolock) ON Memb.Memb_No = Logs.Logs_MembNo"
      vSql = vSQL & " WHERE (Memb.Memb_AcctId = '" & svCustAcctId & "') "
      vSql = vSQL & " AND (Logs.Logs_Type = 'P') "
      If vLearners = "Y" Then
        vSql = vSQL & " AND (Memb.Memb_Level < 3)"
      Else
        vSql = vSQL & " AND (Memb.Memb_Level < 4)"
      End If
      If Len(vStrDate) > 0 Then    
        vSql = vSql & " AND (Logs_Posted > '" & DateAdd("d", -1, vStrDate) & "')"
      End If
      If Len(vEndDate) > 0 Then    
        vSql = vSql & " AND (Logs_Posted < '" & DateAdd("d", 1, vEndDate) & "')"
      End If
      vSql = vSQL & " GROUP BY Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_LastName + ', ' + Memb.Memb_FirstName, Memb.Memb_Id, Memb.Memb_Criteria, LEFT(Logs.Logs_Item, " & vLogItemLength & "), Memb.Memb_Level "
      vSql = vSQL & " ORDER BY Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_Id "

'     sDebug

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

          vId         = Trim(oRs("Id"))
          vModule     = oRs("Module")
          vTimeSpent  = oRs("TimeSpent")
          If vLevel   = "prog" Then
            vTitle    = fProgTitle (Left(vModule, 7))
          Else
            vTitle    = fModsTitle (Right(vModule, 6))
          End If
          vLearner    = Trim(oRs("Learner"))
          If vLearner = "," Then vLearner = ""
    
          '...put a space between different users
          If vId  <> vIdPrev Then 
    %>
    <tr>
      <td valign="top">&nbsp;</td>
      <td valign="top">&nbsp;</td>
      <td valign="top" align="center">&nbsp;</td>
      <td valign="top">&nbsp;</td>
    </tr>
    <tr>
      <td valign="top"><%=fLeft(vLearner, 32)%> </td>
      <td valign="top"><%=fIf(oRs("Memb_Level") < 3, vId, "******")%> </td>
      <td valign="top" align="center">&nbsp;<%=vTimeSpent%> </td>
      <td valign="top"><%=vModule & " - " & fLeft(vTitle, 48)%> </td>
    </tr>
    <%
         Else 
    %>
    <tr>
      <td valign="top">&nbsp;</td>
      <td valign="top">&nbsp;</td>
      <td valign="top" align="center">&nbsp;<%=vTimeSpent%> </td>
      <td valign="top"><%=vModule & " - " & fLeft(vTitle, 48)%> </td>
    </tr>
    <%
          End If
    
          vIDPrev = vID
        
        End If        
        
        
        oRs.MoveNext	        
      Loop
      sCloseDB
    %>
    <tr>
      <td colspan="4" align="center">&nbsp;<p><a href="LogReport0.asp?vStrDate=<%=vStrDate%>&vEndDate=<%=vEndDate%>&vLevel=<%=vLevel%>&vAccount=<%=vAccount%>&vLearners=<%=vLearners%>"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a></p><p>&nbsp;</td>
    </tr>
  </table>
  <%
	  End If
  %>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
