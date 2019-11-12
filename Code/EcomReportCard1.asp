<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="1" style="border-collapse: collapse" bordercolor="#FFFFFF" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td valign="top" colspan="6" align="center">
      <h1 align="center">Ecommerce Report Card</h1>
      <h2 align="left">This report, sorted by Last Name, shows all learners who have purchased content via the ecommerce system&nbsp;and accessed this content between the Start and End Dates.&nbsp; Click on Details to show the Learner's Report Card.&nbsp; Click on the Learner's Name to access the Learner's Profile.</h2></td>
    </tr>
    <tr>
      <th height="30" bgcolor="#F2F9FD" align="left">Learner </th>
      <th height="30" bgcolor="#F2F9FD" align="left">Purchaser</th>
      <th height="30" bgcolor="#F2F9FD" align="left">Purchased</th>
      <th height="30" bgcolor="#F2F9FD" align="left" colspan="3">Program</th>
    </tr>
    <% 
      Dim vNext, vActive, vMods, vModsOnly, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vStrDate, vEndDate, vCredit
      Dim vCurList, vMaxList, vMembIdLast, vBg, vTitle, vLearner, vMemb_Last

      vCurList       = fDefault(Request("vCurList"), 0)
      vMaxList       = fDefault(Request("vMaxList"), 50)
      vStrDate       = Request("vStrDate") 
      vEndDate       = Request("vEndDate") 
      vCredit        = Request("vCredit")
      vFind          = fDefault(Request("vFind"), "S")
      vFindId        = fUnQuote(Request("vFindId"))
      vFindFirstName = fUnQuote(Request("vFindFirstName"))
      vFindLastName  = fUnQuote(Request("vFindLastName"))
      vFindEmail     = fNoQuote(Request("vFindEmail"))

      '...Get initial recordset on first pass and store in session variable
      If vCurList = 0 Then 

        vSql = "SELECT DISTINCT Memb.Memb_No, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Ecom.Ecom_CardName, Ecom.Ecom_Issued, Ecom.Ecom_Programs, V5_Base.dbo.Prog.Prog_Title1 " _

             & " FROM Ecom WITH (nolock) "_
             & "   INNER JOIN Logs WITH (nolock) ON Ecom.Ecom_MembNo = Logs.Logs_MembNo " _
             & "   INNER JOIN Memb WITH (nolock) ON Ecom.Ecom_MembNo = Memb.Memb_No " _
             & "   INNER JOIN V5_Base.dbo.Prog ON Ecom.Ecom_Programs = V5_Base.dbo.Prog.Prog_Id " _

             & " WHERE (Ecom.Ecom_AcctId = '" & svCustAcctId & "') AND (CHARINDEX(Logs.Logs_Type, 'TP') > 0) "_

             & fIf(IsDate(vStrDate), " AND (Logs.Logs_Posted >= '" & vStrDate & "')", "" ) _
             & fIf(IsDate(vEndDate), " AND (Logs.Logs_Posted <= '" & vEndDate & "')", "" ) _

             & fIf(vCredit = "Y", " AND (CHARINDEX('ACCREDITATION', V5_Base.dbo.Prog.Prog_Title1) > 0)", "" ) _

             & fIf(vFind = "S" And Len(vFindId) > 0,        " AND (Memb_Id        LIKE '" & vFindId        & "%') ", "" ) _
             & fIf(vFind = "S" And Len(vFindFirstName) > 0, " AND (Memb_FirstName LIKE '" & vFindFirstName & "%') ", "" ) _
             & fIf(vFind = "S" And Len(vFindLastName) > 0,  " AND (Memb_LastName  LIKE '" & vFindLastName  & "%') ", "" ) _
             & fIf(vFind = "S" And Len(vFindEmail) > 0,     " AND (Memb_Email     LIKE '" & vFindEmail     & "%') ", "" ) _

             & fIf(vFind = "C" And Len(vFindId) > 0,        " AND (Memb_Id        LIKE '%" & vFindId        & "%') ", "" ) _
             & fIf(vFind = "C" And Len(vFindFirstName) > 0, " AND (Memb_FirstName LIKE '%" & vFindFirstName & "%') ", "" ) _
             & fIf(vFind = "C" And Len(vFindLastName) > 0,  " AND (Memb_LastName  LIKE '%" & vFindLastName  & "%') ", "" ) _
             & fIf(vFind = "C" And Len(vFindEmail) > 0,     " AND (Memb_Email     LIKE '%" & vFindEmail     & "%') ", "" ) _

             & " ORDER BY Memb.Memb_LastName, Memb.Memb_FirstName"
 
        sDebug
        sOpenDB
        Set oRs = oDB.Execute(vSql)
        Set Session("soRs") = oRs
        vCurList = 1
      Else  
        Set oRs = Session("soRs")
      End If  

      '...read until either eof or end of group
      Do While Not oRs.Eof

        vMemb_No        = oRs("Memb_No")
        vMemb_Id        = oRs("Memb_Id")
        vMemb_FirstName = oRs("Memb_FirstName")
        vMemb_LastName  = oRs("Memb_LastName")

        vLearner = "********"
        If (svMembLevel > vMemb_Level Or vMemb_No = svMembNo) And InStr(vMemb_Id, vPasswordx) = 0 Then
          vLearner = "<a href='User" & fGroup & ".asp?vMembNo=" & vMemb_No & "&vNext=EcomReportCard_O.asp'>" & vMemb_FirstName & " " & vMemb_LastName & "</a>" & fIf(vMemb_Level = 3, " *", "") & fIf(vMemb_Level = 4, " **", "")
        End If
        
        If vMemb_Last <> vMemb_No Then vBg = "bgcolor='#F2F9FD'" Else vBg = ""
    %>
    <tr>
      <td valign="top" align="left" nowrap height="20" <%=vBg%>><%=vLearner%></td>
      <td valign="top" align="left" height="20" <%=vBg%>><%=fLeft(oRs("Ecom_CardName"), 32)%></td>
      <td valign="top" align="left" nowrap height="20" <%=vBg%>><%=fFormatDate(oRs("Ecom_Issued"))%></td>
      <td valign="top" align="left" height="20" <%=vBg%>><%=oRs("Ecom_Programs")%></td>
      <td valign="top" align="left" height="20" <%=vBg%>><%=fLeft(fClean(oRs("Prog_Title1")), 32)%></td>
      <td valign="top" align="right" height="20" <%=vBg%>>
        <% If vMemb_No <> vMemb_Last Then %>
        <input type="button" onclick="location.href='EcomReportCard2.asp?vMemb_No=<%=vMemb_No%>&vStrDate=<%=vStrDate%>&amp;vEndDate=<%=vEndDate%>&amp;vCurList=<%=vCurList%>&amp;vActive=<%=vActive%>&amp;vFind=<%=vFind%>&amp;vFindId=<%=vFindId%>&amp;vFindFirstName=<%=vFindFirstName%>&amp;vFindLastName=<%=vFindLastName%>&amp;vFindEmail=<%=vFindEmail%>'" value="Details" name="bDetails" class="button085">
        <% End If %>
      </td>
    </tr>
    <%
        vMemb_Last = vMemb_No
        vCurList = vCurList + 1
        If Cint(vCurList) Mod Cint(vMaxList) = 0 Then Exit Do
        oRs.MoveNext
      Loop 
    %>
    <tr>
      <td valign="top" colspan="6" align="center">
      <form method="POST" action="EcomReportCard1.asp">&nbsp;<p><input type="button" onclick="location.href='EcomReportCard.asp?vStrDate=<%=vStrDate%>&amp;vEndDate=<%=vEndDate%>&amp;vCurList=<%=vCurList%>&amp;vActive=<%=vActive%>&amp;vFind=<%=vFind%>&amp;vFindId=<%=vFindId%>&amp;vFindFirstName=<%=vFindFirstName%>&amp;vFindLastName=<%=vFindLastName%>&amp;vFindEmail=<%=vFindEmail%>'" value="Restart" name="bReturn" id="bReturn"class="button085"> <% If Cint(vCurList) > 0 And Cint(vCurList) Mod vMaxList = 0 Then '...If next group, get next starting value %> <%=f10%> <input type="hidden" name="vStrDate" value="<%=vStrDate%>"><input type="hidden" name="vEndDate" value="<%=vEndDate%>"><input type="hidden" name="vCurList" value="<%=vCurList%>"><input type="hidden" name="vFind" value="<%=vFind%>"><input type="hidden" name="vFindId" value="<%=vFindId%>"><input type="hidden" name="vFindFirstName" value="<%=vFindFirstName%>"><input type="hidden" name="vFindLastName" value="<%=vFindLastName%>">
        <input type="hidden" name="vFindEmail" value="<%=vFindEmail%>"><input type="submit" name="bNext" value="Next" class="button085"> <% End If %>
      </p>
      </form></td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  
  <%
     Function fClean(i) '...strip off html tags and notes in brackets
       j = Instr(i, "<")
       If j = 0 Then
         fClean = i
       Else
         fClean = Left(i, j-1)
       End If

       j = Instr(i, "(")
       If j > 1Then
         fClean = Left(fClean, j-1)
       End If

     End Function
  %>

</body>

</html>

