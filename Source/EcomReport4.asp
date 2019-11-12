<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% 
  Dim vStrDate, vEndDate, vStrDateErr, vEndDateErr, vPrograms, vPassword
  
  '...ensure users and/or facilitators don't try to run this report by bypassing the menu page
  If svMembLevel < 4 Then Response.Redirect "Menu.asp"


  Server.ScriptTimeout= 60 * 10
  
  '...defaults to current month
  If Request("vStrDate").Count = 0 And Request("vEndDate").Count = 0 Then
    vStrDate  = Request("vStrDate")      : If Len(vStrDate) = 0 Then vStrDate = fFormatSqlDate(MonthName(Month(Now)) & " 1, " & Year(Now))
    vEndDate  = Request("vEndDate")      : If Len(vEndDate) = 0 Then vEndDate = fFormatSqlDate(DateAdd("d", -1, MonthName(Month(DateAdd("m", +1, Now))) & " 1, " & Year(DateAdd("m", +1, Now))))

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

  vPrograms = fDefault(Request("vPrograms"), "")
  vPassword = fDefault(Request("vPassword"), "")
 

  If Request.Form("bExcel").Count = 1 Then 
    Response.Redirect "EcomReport4X.asp?vPrograms=" & Server.UrlEncode(vPrograms) & "&vPassword=" & Server.UrlEncode(vPassword) & "&vStrDate=" & Server.UrlEncode(vStrDate) & "&vEndDate=" & Server.UrlEncode(vEndDate)
  End If

 

%>
<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

  <% Server.Execute vShellHi %>
  <div align="center">

    <table border="0" cellpadding="4" style="border-collapse: collapse" bordercolor="#DDEEF9">
      <tr>
        <td colspan="7" valign="top" align="center">
        <h1 align="center">Ecommerce Completion Report (Basic)</h1>
        <h2>This report shows the Completion Status of Programs purchased by Individuals, <br>
        sorted by the Learner Name and Purchase Date.</h2>
        <table border="0" cellspacing="0" cellpadding="5" style="border-collapse: collapse" bordercolor="#111111" width="523">
          <form method="POST" action="EcomReport4.asp">
            <input type="Hidden" name="vHidden" value="Hidden">
            <tr>
              <th align="right" valign="top" width="30%" nowrap>Select Start Date :</th>
              <td width="68%" nowrap><input type="text" name="vStrDate" size="9" value="<%=vStrDate%>" class="c2"> <a title="<!--[[-->Start with First Sale.<!--]]-->" onclick="fillField('vStrDate', 'Jan 1, 2000')" href="#">&#937;</a>&nbsp; <span style="background-color: #FFFF00"><%=vStrDateErr%></span><br>
              ie Jan 1, 2010 (MMM DD, YYYY).&nbsp;
              Click &#937; to start with first sale.</td>
            </tr>
            <tr>
              <th align="right" valign="top" width="30%" nowrap>Select End Date :</th>
              <td width="68%" nowrap><input type="text" name="vEndDate" size="9" value="<%=vEndDate%>" class="c2"> <a title="<!--[[-->Finish with Last Sale.<!--]]-->" onclick="fillField('vEndDate', '<%=fFormatDate(DateAdd("d", 1, Now()))%>')" href="#">&#937;</a>&nbsp; <span style="background-color: #FFFF00"><%=vEndDateErr%></span><br>
              ie Mar 31, 2010 (MMM DD, YYYY). 
              Click &#937; to finish with last sale.</td>
            </tr>
            <tr>
              <th align="right" valign="top" width="30%" nowrap>Select Programs :</th>
              <td align="left" width="68%" nowrap><input type="text" name="vPrograms" size="42" value="<%=vPrograms%>" class="c2"> <a title="<!--[[-->Include all Programs<!--]]-->" onclick="fillField('vPrograms', '')" href="#">&#937;</a><br>
              ie P1234EN P2223ES<br>
              Leave empty (&#937;) to show all Programs</td>
            </tr>
            <tr>
              <th align="right" valign="top" width="30%" nowrap>For <%=fIf(svCustPwd, "<!--{{-->Id<!--}}-->", "<!--{{-->Password<!--}}-->")%> :</th>
              <td align="left" width="68%" nowrap><input type="text" name="vPassword" size="20" value="<%=vPassword%>"> <a title="<!--[[-->Include all Programs<!--]]-->" onclick="fillField('vPassword', '')" href="#">&#937;</a><br>Optionally enter ONE <%=fIf(svCustPwd, "<!--{{-->Id<!--}}-->", "<!--{{-->Password<!--}}-->")%> to report on one learner.<br>Leave empty (&#937;) to include all Learners</td>
            </tr>
            <tr>
              <th align="right" valign="top" width="30%" nowrap>Format&nbsp; :</th>
              <td align="left" width="68%" nowrap><input type="submit" value="Online" name="bOnline" class="button070">&nbsp; Maximum 1000 records<br><br> <input type="submit" value="Excel" name="bExcel" class="button070">&nbsp; Maximum 50,000 records<p>&nbsp;</td>
            </tr>
          </form>
        </table>
        </td>
      </tr>

      <% 
        If Request.Form("vHidden").Count > 0 And vStrDateErr = "" And vEndDateErr = "" Then
      %>
      <tr>
        <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Learner</th>
        <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left"><%=fIf(svCustPwd, "<!--{{-->Id<!--}}-->", "<!--{{-->Password<!--}}-->")%></th>
        <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Program</th>
        <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF" align="left">Title</th>
        <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Purchased </th>
        <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Expired</th>
        <th height="20" bgcolor="#DDEEF9" bordercolor="#FFFFFF">Completed</th>
      </tr>
      <% 
        '...for SQL
        Function fProgs(vPrograms)
          fProgs = ""
          If Len(Trim(vPrograms)) > 14 Then
            fProgs = "(Pr.Prog_Id IN ('" & Replace(vPrograms, " ", "', '") & "')) AND "
          Elseif Len(Trim(vPrograms)) = 7 Then
            fProgs = "(Pr.Prog_Id = '" & vPrograms & "') AND "
          End If        
        End Function

        vSql = "SELECT TOP 1000 " _     
             & "  Me.Memb_FirstName + ' ' + Me.Memb_LastName AS Learner, " _ 
             & "  Me.Memb_Id AS Password, " _ 
             & "  Ec.Ecom_Programs AS Program, " _ 
             & "  Pr.Prog_Title1 AS Title,  " _
             & "  Ec.Ecom_Issued AS Purchased, " _ 
             & "  Ec.Ecom_Expires AS Expired, " _ 
             & "  Sc.pcnCompleted AS Completed " _
             & "FROM " _         
             & "  V5_Vubz.dbo.Memb                          AS Me INNER JOIN " _
             & "  V5_Vubz.dbo.Ecom                          AS Ec ON Me.Memb_No = Ec.Ecom_MembNo INNER JOIN " _
             & "  V5_Base.dbo.Prog                          AS Pr ON Ec.Ecom_Programs = Pr.Prog_Id LEFT OUTER JOIN " _
             & "  vuGoldSCORM.dbo.LearnerProgramCompleted   AS Sc ON Me.Memb_No = Sc.pcnMembID AND Pr.Prog_No = Sc.pcnProgramID " _
             & "WHERE "_
             & "  (Ec.Ecom_Media = 'Online') AND "_
             &    fProgs(vPrograms) _
             & "  (Ec.Ecom_Issued BETWEEN '" & vStrDate & "' AND '" & vEndDate & "') AND "_
             & "  (Me.Memb_AcctId = '" & svCustAcctId & "') " _
             &    fIf(vPassword = "", "", " AND Ec.Ecom_Id = '" & vPassword & "'") _
             & "ORDER BY "_
             & "  Me.Memb_LastName, Me.Memb_FirstName, Purchased "
             
'       sDebug
  
        sOpenDb
        Set oRs = oDb.Execute(vSql)
        Do While Not oRs.Eof 
    %>
      <tr>
        <td valign="top" align="left"   nowrap><%=oRs("Learner")%></td>
        <td valign="top" align="left"   nowrap><%=oRs("Password")%></td>
        <td valign="top" align="center" nowrap><%=oRs("Program")%></td>
        <td valign="top" align="left"   nowrap><%=fSmartLeft(oRs("Title"), 40)%></td>
        <td valign="top" align="center" nowrap><%=fFormatDate(oRs("Purchased"))%></td>
        <td valign="top" align="center" nowrap><%=fFormatDate(oRs("Expired"))%></td>
        <td valign="top" align="center" nowrap><%=fFormatDate(oRs("Completed"))%></td>
      </tr>
      <%
          oRs.MoveNext	        
        Loop
        sCloseDB
      End If 
    %>
    </table>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
