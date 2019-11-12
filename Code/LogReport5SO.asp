<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_ExamReport.asp"-->
<!--#include virtual = "V5/Inc/Db_Parm.asp"-->

<%
  Function fId (i)
    fId = fDefault(i, "N/A")
    If oRs("Memb_Level") > 2 Then fId = "******"
  End Function
%> 

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <script src="/V5/Inc/Launch.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <table border="1" width="100%" cellspacing="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9">
    <tr>
      <td valign="top" colspan="6" align="center">
      <h1><!--webbot bot='PurpleText' PREVIEW='Survey Report'--><%=fPhra(000474)%></h1>
      </td>
    </tr>
    <tr>
      <th nowrap height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Group</th>
      <th nowrap height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF"><%=fIf(svCustPwd, fPhraH(000411), fPhraH(000211))%></th>
      <th nowrap height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF"><!--webbot bot='PurpleText' PREVIEW='Learner'--><%=fPhra(000165)%></th>
      <th nowrap height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF"><!--webbot bot='PurpleText' PREVIEW='Date'--><%=fPhra(000112)%></th>
      <th nowrap height="30" align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF"><!--webbot bot='PurpleText' PREVIEW='Title'--><%=fPhra(000019)%> </th>
      <th nowrap height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Results</th>
    </tr>
    <% 
      Dim vCookie, vLevel, vStrDate, vBold, vGrade, vTestExam, vTitle, vCertUrl, vCertType, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindCriteria, vUrl, vResults, vParmNo, aGroup1
      vCookie   = svCustAcctId & "_LogReport5"
      
      vDetails       = Request.Cookies(vCookie)("vDetails") 
      vLevel         = Request.Cookies(vCookie)("vLevel")
      vCurList       = Request.Cookies(vCookie)("vCurList") 
      vStrDate       = Request.Cookies(vCookie)("vStrDate")
      vFind          = Request.Cookies(vCookie)("vFind")
      vFindId        = Request.Cookies(vCookie)("vFindId")
      vFindFirstName = Request.Cookies(vCookie)("vFindFirstName")
      vFindLastName  = Request.Cookies(vCookie)("vFindLastName")
      vFindEmail     = Request.Cookies(vCookie)("vFindEmail")
      vFindCriteria  = Request.Cookies(vCookie)("vFindCriteria")
      vParmNo        = Request.Cookies(vCookie)("vParmNo")


      '...Get initial recordset on first pass and store in session variable
      If vCurList = 0 Then 

        vSql = "SELECT Memb.Memb_No, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Criteria, Memb.Memb_Level, Logs.Logs_Item, Logs_Posted, Memb.Memb_Memo "
        vSql = vSql & " FROM Logs INNER JOIN Memb WITH (nolock) ON Logs_MembNo = Memb_No "

'       vSql = vSql & " WHERE Memb_AcctId= '" & svCustAcctId & "' AND Logs_Type = 'U'"
'       vSql = vSql & " AND Logs.Logs_Posted > '" & vStrDate & "'"
'       vSql = vSql & " AND Memb_Level <= " & svMembLevel

        vSql = vSql & " WHERE Memb_AcctId= '" & svCustAcctId & "'"
        vSql = vSql & " AND Logs.Logs_AcctId = '" & svCustAcctId & "'"
        vSql = vSql & " AND Logs.Logs_Type = 'U'"
        vSql = vSql & " AND Logs.Logs_Posted > '" & vStrDate & "'"
        vSql = vSql & " AND Memb.Memb_Level IN (" & vLevel & ")"


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
      
        '...Group1?
        j = 0
        If Len(vFindCriteria)    > 1 Then 
          aGroup1 = Split(vFindCriteria)
          For i = 0 To Ubound(aGroup1)
            If Cint(aGroup1(i)) > 0 Then
              j = j + 1
              If j = 1 Then 
                vSql = vSql & " AND ((Memb_Criteria LIKE '%" & aGroup1(i) & "%')"
              Else
                vSql = vSql & "  OR (Memb_Criteria LIKE '%" & aGroup1(i) & "%')"
              End If
            End If
          Next
          If j > 0 Then 
             vSql = vSql & " )"
          End If         
        End If

        '...allow a module filter to be extracted from the vParm table via the url [?vParm=2] so report only displays modules required by this user - syntax must be perfect, ie:
'       vSql = vSql & " AND (CHARINDEX(SUBSTRING(Logs.Logs_Item, 9, 4), '0350|0225|0227|0334|0226|0333|0336|0335|0337|0338') > 0) "
        vSql = vSql & " " & fParmValue (vParmNo)

        vSql = vSql & " ORDER BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_No, Memb.Memb_Id, Logs.Logs_Posted "'

'       sDebug
        sOpenDB
        Set oRs = oDB.Execute(vSql)

        Set Session("soRs") = oRs
        vCurList = 1
      '...Else get it from the session variable
      Else  
        Set oRs = Session("soRs")
      End If  

      '...read until either eof or end of group
      Do While Not oRs.Eof
  
        sReadLogsMembSurvey  

        '...contains a program id 
        If Left(vLogs_Item, 1) = "P" Then  
          vLogs_Module = Mid(vLogs_Item, 9, 6)
          vResults = Mid(vLogs_Item, 16) 
        '...contains an 'undefined' program id 
        Else
          vLogs_Module = Mid(vLogs_Item, 11, 6)
          vResults = Mid(vLogs_Item, 18) 
        End If

        vTitle   = fModsTitle(vLogs_Module)  

        '...ensure you can only see members with same criteria
'       If fCriteriaOk (svMembCriteria, vMemb_Criteria) Then
          vCurList = vCurList + 1
    %>
    <tr>
      <td valign="top" align="left"><%=Replace(fCriteria(vMemb_Criteria), "+", "<br>")%></td>
      <td valign="top" nowrap><%=fId(vMemb_Id)%>&nbsp; </td>
      <td valign="top" nowrap>
        <% If svMembLevel > 3 Then %>
        <a title="Vubiz Learner No: <%=vMemb_No%>" href="#"><%=fLeft(vMemb_FirstName & " " & vMemb_LastName, 24)%></a>
        <% Else %>
        <%=fLeft(vMemb_FirstName & " " & vMemb_LastName, 24)%>
        <% End If %>
      </td>
      <td valign="top" align="center" nowrap><%=fFormatDate (vLogs_Posted)%></td>
      <td valign="top" nowrap><%=vLogs_Module & " - " & vTitle%></td>
      <td valign="top"><%=vResults%> </td>
    </tr>
    <%
        'End If

        oRs.MoveNext
        If Cint(vCurList) Mod 50 = 0 Then Exit Do
      Loop 
    %>      

    <tr>
      <td bgcolor="#FFFFFF" align="center" colspan="7" height="70" class="c2">
      <%
        '...If next group
        If Cint(vCurList) > 0 And Cint(vCurList) Mod 50 = 0 Then
          Response.Cookies(vCookie)("vCurList") = vCurList
      %>
      <a href="LogReport5SO.asp"><!--webbot bot='PurpleText' PREVIEW='Next Group'--><%=fPhra(000834)%></a>
      <%
        Else 
          Set oRs = Nothing
        End If      
      %>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
      <a href="LogReport5.asp"><!--webbot bot='PurpleText' PREVIEW='Restart Report'--><%=fPhra(000225)%></a></td>
    </tr>

  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

