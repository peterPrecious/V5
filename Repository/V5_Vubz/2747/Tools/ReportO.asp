<!--#include virtual = "V5\Inc\Setup.asp"-->
<!--#include virtual = "V5\Inc\Initialize.asp"-->
<!--#include virtual = "V5\Inc\Db_Cust.asp"-->
<!--#include virtual = "V5\Inc\Db_Memb.asp"-->
<!--#include virtual = "V5\Inc\Db_Mods.asp"-->

<%
  Function fId (vId, vLevel)
    fId = fIf(vLevel > 2, "******", fDefault(vId, "N/A"))
  End Function
%>

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script language="JavaScript" src="/V5/Inc/Launch.js"></script>
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <title></title>
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <table border="1" width="100%" cellspacing="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9">
    <tr>
      <td valign="top" colspan="10" align="center"><h1>Assessment Report</h1></td>
    </tr>
    <tr>
      <th height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Learner</th>
      <th height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Email</th>
      <th height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Student Id</th>
      <th height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Institution</th>
      <th height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Faculty</th>
      <th height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Course</th>
      <th height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Academic Year</th>
      <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Date</th>
      <th height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Title </th>
      <th height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Score | Result</th>
    </tr>
    <% 
      Dim vStrDate, vBold, vGrade, vTestExam, vTitle, vCertUrl, vCertType, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vUrl, aMemo
      Dim vLogs_No, vLogs_AcctId, vLogs_Type, vLogs_Item, vLogs_Posted, vLogs_MembNo
      Dim vLogs_Module, vLogs_Result, vDetails, vCurList, vMaxList
      
      vDetails       = Request("vDetails") 
      vCurList       = Request("vCurList") 
      vStrDate       = Request("vStrDate")
      vFind          = fDefault(Request("vFind"), "S")
      vFindId        = fUnQuote(Request("vFindId"))
      vFindFirstName = fUnQuote(Request("vFindFirstName"))
      vFindLastName  = fUnQuote(Request("vFindLastName"))
      vFindEmail     = fNoQuote(Request("vFindEmail"))

      '...Get initial recordset on first pass and store in session variable
      If vCurList = 0 Then 

        vSql = "SELECT Memb.Memb_No, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Memo, Memb.Memb_Level "
        If vDetails = "y" Then  '...details of assessments
          vSql = vSql & ", Left(Logs.Logs_Item, 6) AS Logs_Module, CAST(Right(Logs.Logs_Item, 3) AS FLOAT) AS Logs_Result, Logs.Logs_Posted "
        ElseIf vDetails = "n" Then '...summary
          vSql = vSql & ", Left(Logs.Logs_Item, 6) AS Logs_Module,  MAX(CAST(Right(Logs.Logs_Item, 3) AS FLOAT)) AS Logs_Result, MAX(Logs.Logs_Posted) AS Logs_Posted "
        ElseIf vDetails = "s" Then '...details of surveys
          vSql = vSql & ", SUBSTRING(Logs.Logs_Item, 9, 6) AS Logs_Module,  SUBSTRING(Logs.Logs_Item, 16, 999) AS Logs_Result, Logs.Logs_Posted "
        End If

        vSql = vSql & " FROM Logs WITH (nolock) INNER JOIN Memb WITH (nolock) ON Logs_MembNo = Memb_No "
        vSql = vSql & " WHERE Logs_AcctId= '" & svCustAcctId & "' AND Logs_Type = '" & fIf(vDetails = "s", "U", "T") & "'"
        vSql = vSql & " AND Logs.Logs_Posted > '" & vStrDate & "'"
        vSql = vSql & " AND Memb_Level < 4 "

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
      
        If vDetails = "y" or vDetails = "s" Then    
          vSql = vSql & " ORDER BY Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_No, Memb.Memb_Id, Memb.Memb_Memo, Logs.Logs_Posted "'
        Else
          vSql = vSql & " GROUP BY Memb.Memb_LastName, Memb.Memb_FirstName, Left(Logs.Logs_Item, 6), Memb.Memb_No, Memb.Memb_Id, Memb.Memb_Memo, Memb.Memb_Level"
          vSql = vSql & " ORDER BY Memb.Memb_LastName, Memb.Memb_FirstName, Left(Logs.Logs_Item, 6), Memb.Memb_No, Memb.Memb_Id, Memb.Memb_Memo "
        End If

'       sDebug "vSql", vSql

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
  
        vLogs_Module                = oRs("Logs_Module")
        vLogs_Result                = oRs("Logs_Result")
        vLogs_Posted                = oRs("Logs_Posted")

        vMemb_Level                 = oRs("Memb_Level")
        vMemb_No                    = oRs("Memb_No")
        vMemb_Id                    = oRs("Memb_Id")
        vMemb_Id                    = fId(vMemb_Id, vMemb_Level)

        vMemb_FirstName             = oRs("Memb_FirstName")
        vMemb_LastName              = oRs("Memb_LastName")
        vMemb_Memo									= oRs("Memb_Memo")

        aMemo                       = Split(vMemb_Memo, "|")
        If Ubound(aMemo) < 5 Then 
          vMemb_Memo                = vMemb_Memo & "||||"
          aMemo                     = Split(vMemb_Memo, "|")
        End If
        
        vTitle = fModsTitle(vLogs_Module)

        vCurList = vCurList + 1


    %>
    <tr>
      <td valign="top" nowrap><%=fLeft(vMemb_FirstName & " " & vMemb_LastName, 24)%> </td>
      <td valign="top" nowrap><%=vMemb_Id%></td>
      <td valign="top"><%=fLeft(aMemo(0), 24)%></td>
      <td valign="top"><%=fLeft(aMemo(1), 24)%></td>
      <td valign="top"><%=fLeft(aMemo(2), 24)%></td>
      <td valign="top"><%=fLeft(aMemo(4), 24)%></td>
      <td valign="top"><%=fLeft(aMemo(3), 24)%></td>
      <td valign="top" align="center" nowrap><%=fFormatDate (vLogs_Posted)%></td>
      <td valign="top"><%=fIf(svMembLevel > 3, vLogs_Module & " - ", "") & fLeft(vTitle, 24)%></td>
      <td valign="top" align="left"><%=vLogs_Result%></td>
    </tr>
    <%

        oRs.MoveNext
        If Cint(vCurList) Mod 50 = 0 Then Exit Do
      Loop 
    %>
    <tr>
      <td bgcolor="#FFFFFF" valign="top" align="center" colspan="10"><p>&nbsp;</p><%

      '...If next group
      If Cint(vCurList) > 0 And Cint(vCurList) Mod 50 = 0 Then
    %>
      <form method="POST" action="ReportO.asp">
        <p><input type="hidden" name="vCurList" value="<%=vCurList%>"><input type="submit" value="Next Group" name="bNext" class="button"></p>
      </form>
    <%
      Else 
        Set oRs = Nothing
      End If
      
      vUrl = "Report.asp" _
           & "?vStrDate="       & Server.UrlEncode(vStrDate)       _
           & "&vDetails="       & vDetails       _
           & "&vCurList="       & vCurList       _
           & "&vFind="          & vFind          _
           & "&vFindId="        & vFindId        _
           & "&vFindFirstName=" & vFindFirstName _
           & "&vFindLastName="  & vFindLastName  _
           & "&vFindEmail="     & vFindEmail
    %>
    <h2><a href="<%=vUrl%>"><!--[[-->Restart Report<!--]]--></a></h2><p>&nbsp;</p></td>
    </tr>
  </table>
  <!--#include virtual = "V5\Inc\Shell_Lo.asp"-->

</body>

</html>
