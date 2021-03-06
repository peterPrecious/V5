<!--#include virtual = "V5\Inc\Setup.asp"-->
<!--#include virtual = "V5\Inc\Initialize.asp"-->
<!--#include virtual = "V5\Inc\Db_Cust.asp"-->
<!--#include virtual = "V5\Inc\Db_Memb.asp"-->
<!--#include virtual = "V5\Inc\Db_Mods.asp"-->
<!--#include virtual = "V5\Inc\Db_Crit.asp"-->

<%
    Dim vStrDate, vBold, vGrade, vTestExam, vTitle, vCertUrl, vCertType, vFind, vFindId, vFindFirstName, vFindLastName, vFindCriteria, vFindEmail, vUrl, aMemo, aCrit
    Dim vLogs_No, vLogs_AcctId, vLogs_Type, vLogs_Item, vLogs_Posted, vLogs_MembNo
    Dim vLogs_Module, vScore, vCurList, vMaxList, vSum, vCriteria
    
    vCurList        = Request("vCurList") 
    vStrDate        = Request("vStrDate")
    vFind           = fDefault(Request("vFind"), "S")
    vFindId         = fUnQuote(Request("vFindId"))
    vFindFirstName  = fUnQuote(Request("vFindFirstName"))
    vFindLastName   = fUnQuote(Request("vFindLastName"))
    vFindCriteria   = Request("vFindCriteria") 
    vFindEmail      = fNoQuote(Request("vFindEmail"))
%>

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <title></title>
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <table border="1" width="100%" cellspacing="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9">
    <tr>
      <td valign="top" colspan="8" align="center">
        <h1><br>Passport to Safety Report Card</h1>
        <p>Ordered by Institution/Faculty/Name</p>
        <h6>Note: only Learners who have completed all four assessments appear on this Report.</h6>
      </td>
    </tr>
    <tr>
      <th height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Name</th>
      <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Final Score<br>on Tests</th>
      <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Student #</th>
      <th height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Institution</th>
      <th height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Faculty</th>
      <th height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Course</th>
      <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Date <br>Completed</th>
      <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Email</th>
    </tr>

    <%

      '...Get initial recordset on first pass and store in session variable
      If vCurList = 0 Then 

        vSql = "SELECT Memb.Memb_No, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Memo, Memb.Memb_Criteria, " 
        vSql = vSql & " AVG(CAST(RIGHT(Logs.Logs_Item, 3) AS FLOAT)) AS [Score], MAX(Logs.Logs_Posted) AS Logs_Posted, SUM(1) AS [Sum]"
        vSql = vSql & " FROM Logs WITH (nolock) INNER JOIN Memb WITH (nolock) ON Logs_MembNo = Memb_No "
        vSql = vSql & " WHERE (Logs_AcctId= '" & svCustAcctId & "') AND (Logs_Type = 'T')"

		    If vFindCriteria <> "0" Then
          vSql = vSql & " AND (CHARINDEX(Memb.Memb_Criteria, '" & vFindCriteria & "') > 0)"
        End If

        vSql = vSql & " AND (CHARINDEX(LEFT(Logs.Logs_Item, 4), '9427 9495 9497 9498') > 0) "
        vSql = vSql & " AND (Logs.Logs_Posted > '" & vStrDate & "')"
        vSql = vSql & " AND (Memb_Level = 2) "
        vSql = vSql & " AND (Memb.Memb_Criteria <> '0')"

        If vFind = "S" Then
          If Len(vFindId)        > 0 Then vSql = vSql & " AND (Memb_Id        LIKE '" & vFindId         & "%')"
          If Len(vFindFirstName) > 0 Then vSql = vSql & " AND (Memb_FirstName LIKE '" & vFindFirstName  & "%')"
          If Len(vFindLastName)  > 0 Then vSql = vSql & " AND (Memb_LastName  LIKE '" & vFindLastName   & "%')"
          If Len(vFindEmail)     > 0 Then vSql = vSql & " AND (Memb_Email     LIKE '" & vFindEmail      & "%')"
        Else
          If Len(vFindId)        > 0 Then vSql = vSql & " AND (Memb_Id        LIKE '%" & vFindId        & "%')"
          If Len(vFindFirstName) > 0 Then vSql = vSql & " AND (Memb_FirstName LIKE '%" & vFindFirstName & "%')"
          If Len(vFindLastName)  > 0 Then vSql = vSql & " AND (Memb_LastName  LIKE '%" & vFindLastName  & "%')"
          If Len(vFindEmail)     > 0 Then vSql = vSql & " AND (Memb_Email     LIKE '%" & vFindEmail     & "%')"
        End If
      
        vSql = vSql & " GROUP BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_No, Memb.Memb_Id, Memb.Memb_Memo "
        vSql = vSql & " ORDER BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_No, Memb.Memb_Id, Memb.Memb_Memo "

'       sDebug

        sOpenDb
        Set oRs = oDb.Execute(vSql)

        Set Session("soRs") = oRs
        vCurList = 1

      '...Else get it from the session variable
      Else  

        Set oRs = Session("soRs")
      End If  


      '...read until either eof or end of group
      Do While Not oRs.Eof
  
        vScore                      = oRs("Score")
        vLogs_Posted                = oRs("Logs_Posted")
        vSum                        = oRs("Sum")

        vMemb_No                    = oRs("Memb_No")
        vMemb_Id                    = oRs("Memb_Id")

        vMemb_FirstName             = oRs("Memb_FirstName")
        vMemb_LastName              = oRs("Memb_LastName")

        vMemb_Criteria   			      = oRs("Memb_Criteria")
    	  vCriteria				            = fCriteria (vMemb_Criteria)
    	  aCrit					              = Split(vCriteria , "|")
        If Ubound(aCrit) <> 1 Then 
          aCrit                     = Split("||", "|")
        End If
    
        vMemb_Memo					        = fOkValue(oRs("Memb_Memo"))
        aMemo                       = Split(vMemb_Memo, "|")
        If Ubound(aMemo) < 5 Then 
          vMemb_Memo                = vMemb_Memo & "||||"
          aMemo                     = Split(vMemb_Memo, "|")
        End If
       
        vSum                        = Cint(oRs("Sum"))

        If vSum = 4 Then 
          vCurList = vCurList + 1


    %>
    <tr>
      <td valign="top" nowrap><%=fLeft(vMemb_FirstName & " " & vMemb_LastName, 24)%> </td>
      <td valign="top" nowrap align="center"><%=vScore%></td>
      <td valign="top" align="left"><%=fLeft(aMemo(0), 24)%></td>
      <td valign="top"><%=fLeft(aCrit(0), 24)%></td>
      <td valign="top"><%=fLeft(aCrit(1), 24)%></td>
      <td valign="top"><%=fLeft(aMemo(4), 24)%></td>
      <td valign="top" align="center" nowrap><%=fFormatDate (vLogs_Posted)%></td>
      <td valign="top" align="left"><%=vMemb_Id%></td>
    </tr>
    <%
        End If
        oRs.MoveNext
        If Cint(vCurList) Mod 100 = 0 Then Exit Do
      Loop 
    %>
    <tr>
      <td bgcolor="#FFFFFF" valign="top" align="center" colspan="8"><p>&nbsp;</p>
      
    <%
      '...If next group
      If Cint(vCurList) > 0 And Cint(vCurList) Mod 100 = 0 Then
    %>

      <form method="POST" action="ReportCard_1O.asp">
        <p><input type="hidden" name="vCurList" value="<%=vCurList%>">
        <input type="submit" value="Next Group" name="bNext" class="button"></p>
      </form>

    <%
      Else 
        Set oRs = Nothing
      End If
      
      vUrl = "ReportCard_1.asp" _
           & "?vStrDate="        & Server.UrlEncode(vStrDate) _
           & "&vFindCriteria="   & vFindCriteria   _
           & "&vCurList="        & vCurList        _
           & "&vFind="           & vFind           _
           & "&vFindId="         & vFindId         _
           & "&vFindFirstName="  & vFindFirstName  _
           & "&vFindLastName="   & vFindLastName   _
           & "&vFindEmail="      & vFindEmail
    %>

    <h2><a href="<%=vUrl%>"><!--[[-->Restart Report<!--]]--></a></h2><p>&nbsp;</p></td>
    </tr>
  </table>
  <!--#include virtual = "V5\Inc\Shell_Lo.asp"-->

</body>

</html>
