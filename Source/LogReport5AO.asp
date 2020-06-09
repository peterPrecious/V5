<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_ExamReport.asp"-->
<!--#include file = "ModuleStatusRoutines.asp"-->
<!--#include virtual = "V5/Inc/Db_Parm.asp"-->

<%
  Function fId (i)
    fId = fDefault(i, "N/A")
    If oRs("Memb_Level") >= svMembLevel  Then fId = "******"
  End Function
  
  sGetCust(svCustId)
%> 

<html>

<head>
  <title>LogReport5AO - Assessment Report</title>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <script src="/V5/Inc/Launch.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %> 
</head>

<body>

  <% Server.Execute vShellHi %>

  <h1><!--[[-->Assessment Report<!--]]--></h1>
  <h2>This is the Basic Assessment Report.&nbsp; For a more advanced version, use the Learner Report Card.</h2>
  <br />


  <table class="table">
    <tr>
      <th class="rowshade" style="width:10%">Group</th>
      <th class="rowshade" style="width:10%"><%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%></th>
      <th class="rowshade" style="width:15%"><!--[[-->Learner<!--]]--></th>
      <th class="rowshade" style="width:10%"><!--[[-->Date<!--]]--></th>
      <th class="rowshade" style="width:40%"><!--[[-->Title<!--]]--> </th>
      <th class="rowshade" style="width:05%"><!--[[-->Score<!--]]--></th>
    </tr>
    <% 
      Dim vCookie, vLevel, vScore1, vScore2, vStrDate, vBold, vGrade, vTestExam, vTitle, vCertUrl, vCertType
      Dim vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vUrl, vResults, vParmNo, aGroup1
      vCookie   = svCustAcctId & "_LogReport5"

      vDetails       = Request.Cookies(vCookie)("vDetails") 
      vLevel         = Request.Cookies(vCookie)("vLevel")
      vScore1        = Request.Cookies(vCookie)("vScore1")
      vScore2        = Request.Cookies(vCookie)("vScore2")
      vCurList       = Request.Cookies(vCookie)("vCurList") 
      vStrDate       = Request.Cookies(vCookie)("vStrDate")
      vFind          = Request.Cookies(vCookie)("vFind")
      vFindId        = Request.Cookies(vCookie)("vFindId")
      vFindFirstName = Request.Cookies(vCookie)("vFindFirstName")
      vFindLastName  = Request.Cookies(vCookie)("vFindLastName")
      vFindEmail     = Request.Cookies(vCookie)("vFindEmail")
      vFindMemo      = Request.Cookies(vCookie)("vFindMemo")
      vFindCriteria  = Request.Cookies(vCookie)("vFindCriteria")
      vParmNo        = Request.Cookies(vCookie)("vParmNo")

      '...old values before long mods
      ' CASE LEN(Logs_Item) WHEN 10 THEN 'M' ELSE 'E' END 


      '...Get initial recordset on first pass and store in session variable
      If vCurList = 0 Then 

        vSql = "SELECT Memb.Memb_No, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Criteria, Memb.Memb_Level, Memb.Memb_Memo "
        If vDetails = "y" Then    
          vSql = vSql & ",  Left(Logs.Logs_Item, CHARINDEX('_', Logs_Item) - 1) AS Logs_Module, CAST(Right(Logs.Logs_Item, 3) AS FLOAT) AS Logs_Grade, Logs.Logs_Posted, CASE CHARINDEX('_', Logs_Item, CHARINDEX('_', Logs_Item) + 1) WHEN 0 THEN 'M' ELSE 'E' END AS [Logs_Assess] "
        Else
          vSql = vSql & ",  Left(Logs.Logs_Item, CHARINDEX('_', Logs_Item) - 1) AS Logs_Module, MAX(CAST(Right(Logs.Logs_Item, 3) AS FLOAT)) AS Logs_Grade, MAX(Logs.Logs_Posted) AS Logs_Posted, CASE CHARINDEX('_', Logs_Item, CHARINDEX('_', Logs_Item) + 1) WHEN 0 THEN 'M' ELSE 'E' END AS [Logs_Assess] "
        End If
        vSql = vSql & " FROM Logs INNER JOIN Memb WITH (nolock) ON Logs_MembNo = Memb_No "_
                    & " WHERE Memb_AcctId= '" & svCustAcctId & "'"_
                    & " AND Logs.Logs_AcctId = '" & svCustAcctId & "'"_
                    & " AND Logs.Logs_Type = 'T'"_
                    & " AND Logs.Logs_Posted > '" & vStrDate & "'"_
                    & " AND Memb.Memb_Level IN (" & vLevel & ")"_
                    & " AND CAST(Right(Logs.Logs_Item, 3) AS FLOAT) " & fIf(vScore1 = "GE", ">= ", "<= ") & vScore2

        If vFind = "S" Then
          If Len(vFindId)        > 0 Then vSql = vSql & " AND (Memb_Id        LIKE '" & vFindId         & "%')"
          If Len(vFindFirstName) > 0 Then vSql = vSql & " AND (Memb_FirstName LIKE '" & vFindFirstName  & "%')"
          If Len(vFindLastName)  > 0 Then vSql = vSql & " AND (Memb_LastName  LIKE '" & vFindLastName   & "%')"
          If Len(vFindEmail)     > 0 Then vSql = vSql & " AND (Memb_Email     LIKE '" & vFindEmail      & "%')"
          If Len(vFindMemo)      > 0 Then vSql = vSql & " AND (Memb_Memo      LIKE '" & vFindMemo       & "%')"
        Else
          If Len(vFindId)        > 0 Then vSql = vSql & " AND (Memb_Id        LIKE '%" & vFindId        & "%')"
          If Len(vFindFirstName) > 0 Then vSql = vSql & " AND (Memb_FirstName LIKE '%" & vFindFirstName & "%')"
          If Len(vFindLastName)  > 0 Then vSql = vSql & " AND (Memb_LastName  LIKE '%" & vFindLastName  & "%')"
          If Len(vFindEmail)     > 0 Then vSql = vSql & " AND (Memb_Email     LIKE '%" & vFindEmail     & "%')"
          If Len(vFindMemo)      > 0 Then vSql = vSql & " AND (Memb_Memo      LIKE '%" & vFindMemo      & "%')"
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
        vSql = vSql & " " & fParmValue (vParmNo)

        If vDetails = "y" Then    
          vSql = vSql & " ORDER BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_No, Memb.Memb_Id, Logs.Logs_Posted "'
        Else
          vSql = vSql & " GROUP BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Left(Logs.Logs_Item, CHARINDEX('_', Logs_Item) - 1), Memb.Memb_No, Memb.Memb_Id, Memb.Memb_Level, CASE CHARINDEX('_', Logs_Item, CHARINDEX('_', Logs_Item) + 1) WHEN 0 THEN 'M' ELSE 'E' END, Memb.Memb_Memo "
          vSql = vSql & " ORDER BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Left(Logs.Logs_Item, CHARINDEX('_', Logs_Item) - 1), Memb.Memb_No, Memb.Memb_Id "
        End If

'       sDebug

        sOpenDb
        Set oRs = oDB.Execute(vSql)

        Set Session("soRs") = oRs
        vCurList = 1
      '...Else get it from the session variable
      Else  
        Set oRs = Session("soRs")
      End If  

      '...read until either eof or end of group


      'stop
      Do While Not oRs.Eof
  
        sReadLogsMemb      

        '...ensure you can only see members with same criteria
'       If fCriteriaOk (svMembCriteria, vMemb_Criteria) Then
          vCurList = vCurList + 1

          '...get title
          If vLogs_Assess = "E" Then
            vCertType = "Exam"
            vTitle = fExamTitle(vLogs_Module)
            Session("CertProg") = fProgCert (vLogs_Module)  '...Flag ProgCerts
            vCertUrl = "javascript:jCertificate('" & svLang & "','" & vLogs_Module & "','" & fjUnquote(vTitle) & "','" & vLogs_Posted & "','" & vLogs_Grade/100 & "','Exam', '" & fjUnquote(vMemb_FirstName) & " " & fjUnquote(vMemb_LastName) & "')"
          Else  
            vCertType = "Test"
            vTitle = fModsTitle(vLogs_Module)
            '...either platform or vuassess
            If fModsVuCert (vLogs_Module) Then 
              vCertUrl = "javascript:jVuCertificate('" & vLogs_Module & "','" & fjUnquote(vTitle) & "','" & fFormatDate(vLogs_Posted) & "','" & vLogs_Grade/100 & "','" & fjUnquote(vMemb_FirstName) & "','" & fjUnquote(vMemb_LastName) & "','" & svCustBanner & "')"
            Else
              Session("CertProg") = fProgCert (vLogs_Module)  '...Flag ProgCerts
              vCertUrl = "javascript:jCertificate('" & svLang & "','" & vLogs_Module & "','" & fjUnquote(vTitle) & "','" & vLogs_Posted & "','" & vLogs_Grade/100 & "','Test', '" & fjUnquote(vMemb_FirstName) & " " & fjUnquote(vMemb_LastName) & "')"
            End If
          End If
    %>
    <tr>
      <td><%=Replace(fCriteria(vMemb_Criteria), "+", "<br>")%></td>
      <td><%=fId(vMemb_Id)%>&nbsp; </td>
      <td>
        <% If svMembLevel > 3 Then %>
        <a title="Vubiz Learner No: <%=vMemb_No%>" href="#"><%=fLeft(vMemb_FirstName & " " & vMemb_LastName, 24)%></a>
        <% Else %>
        <%=fLeft(vMemb_FirstName & " " & vMemb_LastName, 24)%>
        <% End If %>
      </td>
      <td style="text-align:center"><%=fFormatDate (vLogs_Posted)%></td>
      <td><%=vLogs_Module & " - " & vTitle%></td>
      <td style="text-align:center"><%=vLogs_Grade%></td>
    </tr>
    <%
        oRs.MoveNext
        If Cint(vCurList) Mod 50 = 0 Then Exit Do
      Loop 
    %>      

    <tr>
      <td colspan="7" style="text-align:center; padding:20px;">
      <%
        '...If next group
        If Cint(vCurList) > 0 And Cint(vCurList) Mod 50 = 0 Then
          Response.Cookies(vCookie)("vCurList") = vCurList
      %>
      <a href="LogReport5AO.asp"><!--[[-->Next Group<!--]]--></a>
      <%
        Else 
          Set oRs = Nothing
        End If      
      %><%=f10%>
      <a href="LogReport5.asp"><!--[[-->Restart Report<!--]]--></a></td>
    </tr>

  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>