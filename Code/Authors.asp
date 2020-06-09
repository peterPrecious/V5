<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vActive, vGlobal, vCustId, vNext, vEdit, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vFormat, vLearners, vLevel
  Dim vLastValue, vDetails, vCurList, vRecnt, vWhere, aCrit, bG2CustEmail, bG2MembEmail, vCols


%>

<html>

<head>
  <title>Learner Report</title>
  <meta charset="UTF-8">
  <% If vRightClickOff Then %><script JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">


</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="1" width="100%" cellspacing="0" bordercolor="#DDEEF9" cellpadding="3" style="border-collapse: collapse">
    <tr>
      <td valign="top" colspan="7" align="center"><h1>Authors</h1></td>
    </tr>
    <tr>
      <th align="left" bgcolor="#DDEEF9" height="30%" nowrap bordercolor="#FFFFFF">Account</th>
      <th align="left" bgcolor="#DDEEF9" height="30%" nowrap bordercolor="#FFFFFF">First Name</th>
      <th align="left" bgcolor="#DDEEF9" height="30%" nowrap bordercolor="#FFFFFF">Last Name</th>
      <th bgcolor="#DDEEF9" height="30%" nowrap bordercolor="#FFFFFF">Active</th>
      <th align="left" bgcolor="#DDEEF9" height="30%" nowrap bordercolor="#FFFFFF">Level</th>
      <th bgcolor="#DDEEF9" height="30%" nowrap bordercolor="#FFFFFF">Memb_Mgr</th>
      <th bgcolor="#DDEEF9" height="30%" nowrap bordercolor="#FFFFFF">Memb_LCMS</th>
      <th bgcolor="#DDEEF9" height="30%" nowrap bordercolor="#FFFFFF">Memb_Author</th>
      <th bgcolor="#DDEEF9" height="30%" nowrap bordercolor="#FFFFFF">Email</th>
    </tr>
    <tr>
      <td valign="top" colspan="<%=vCols%>">&nbsp;</td>
    </tr>
    <%  

      vSql = ""_
           & "  SELECT Cust.Cust_Id, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Level, Memb.Memb_Active, Memb.Memb_Auth, Memb.Memb_VuBuild, Memb.Memb_Manager, Memb_LCMS
           & "  FROM Cust INNER JOIN "_
           & "  Memb ON Cust.Cust_AcctId = Memb.Memb_AcctId
           & "  WHERE     (Cust.Cust_Level = 7) AND (Memb.Memb_Internal = 0)
           & "  ORDER BY Memb.Memb_Id

      Do While Not oRs.Eof
        sReadMemb
  
        '...determine if this is a G2 learner (bG2MembEmail) for Resend Button
        bG2MembEmail = fIf(bG2CustEmail And vMemb_EcomG2alert And vMemb_Level = 2, True, False)

        vMemb_Id = fDefault(vMemb_Id, "N/A")

        '...ensure you can only see users below your level
        j = ""
        If vMemb_Level = 3 Then
          j = "<b> * </b>"
        ElseIf vMemb_Level = 4 Then
          j = "<b> ** </b>"
        ElseIf vMemb_Level = 5 Then
          j = "<b> *** </b>"
        End If     

        '...display if sponsored
        k = ""
        If vMemb_Sponsor > 0 Then
          k = "(<a href='User" & fGroup & ".asp?vMembNo=" & vMemb_Sponsor & "&vNext=" & vNext & "'>" & fPhraH(000492) & "</a>)"
        End If

        vCurList = vCurList + 1
    %>
    <tr>
      <td valign="top" nowrap align="left">
        <%=fIf(Len(Trim(vMemb_Criteria)) < 3 Or Trim(vMemb_Criteria) = "0" , "", Replace(fCriteria(vMemb_Criteria), " + ", "<br>"))%>
        <%=fIf(vMemb_Group2 = 0 , "", "  [" & vMemb_Group2 & "]")%>
        <% If svMembLevel = 5 Then %><br><font color="#3977B6"><%=fRights()%></font><% End If %>
      </td>
      <td valign="top" nowrap align="left">
        <%=fLeft(vMemb_FirstName, 16) & " " & fLeft(vMemb_LastName, 16) & fIf(Len(vMemb_Organization) > 0, ", " & vMemb_Organization, "") & "<br><font color='#3977B6'>" & vMemb_Email & "</font>"%> 
      </td>
      <td valign="top" nowrap align="left">
        &nbsp;</td>
      <td align="center" valign="top" nowrap>
      <% If svMembLevel < 5 Then %>
      <%=fYN(vMemb_Active)%>
      <% Else %>      
      <a href="#" onclick="history(<%=vMemb_No%>)"><%=fYN(vMemb_Active)%></a>
			<% End If %>
      </td>
      <td valign="top" nowrap>
        <% If (svMembManager Or svMembLevel > vMemb_Level Or vMemb_No = svMembNo) And InStr(vMemb_Id, vPasswordx) = 0 Then %>  
          <%=fGlobal%><a href="<%=vEdit%>?vMembNo=<%=vMemb_No%>&vCustId=<%=vCustId%>&vNext=<%=vNext%>"><%=vMemb_Id%></a> <%=j%> <%=k%>
        <% Else %>********&nbsp; 
        <% End If %> 
        <br><font color="#3977B6"><%=vMemb_Memo%></font> 
      </td>
      <td align="center" valign="top" nowrap><%=fFormatDate(vMemb_FirstVisit)%><br><font color="#3977B6"><%=fFormatDate(vMemb_LastVisit)%></font> </td>
      <td align="center" valign="top" nowrap><%=fFormatDate(fIf(bG2CustEmail, vCust_Expires, vMemb_Expires))%></td>
      <td align="center" valign="top" nowrap>&nbsp;<%=vMemb_NoVisits & "<br><font color='#3977B6'>" & FormatNumber(vMemb_NoHours/60,1) & "</font>"%> </td>
      <% If bG2CustEmail Then %>
      <td align="center" valign="top" nowrap>      
        <% If bG2MembEmail And Len(Trim(vMemb_Programs)) > 0 Then %>
        <input onclick="resendEmails(<%=vMemb_No%>, '<%=svLang%>')" type="button" value="Resend" name="bResend" class="button">      
        <% Else %>&nbsp;
        <% End If %>     
      </td>
      <% End If %>     
    </tr>
    <%
        oRs.MoveNext
        If Cint(vCurList) > 0 And Cint(vCurList) Mod 50 = 0 Then Exit Do
      Loop
      Set oRs = Nothing
      sCloseDb
    %>
    
  </table>


  <div align="center">
    <form method="POST" action="Users_o.asp">


    <table cellspacing="0" cellpadding="10" border="0">
      <tr>
        <% If Len(vNext) > 0 Then %>
        <td bgcolor="#FFFFFF" align="right">
          <input type="button" onclick="location.href='<%=vNext%>'" value="<%=bReturn%>" name="bReturn" id="bReturn"class="button085">
        </td>
        <% End If %>
        <td bgcolor="#FFFFFF" align="right">
          <input type="button" onclick="location.href='Users.asp?vGlobal=<%=vGlobal%>&vNext=<%=vNext%>&vLearners=<%=vLearners%>&vCustId=<%=vCustId%>'" value="<%=bRestart%>" name="bRestart" class="button085">
        </td>

        <% If Cint(vCurList) > 0 And Cint(vCurList) Mod 50 = 0 Then '...If next group, get next starting value %>
        <td bgcolor="#FFFFFF" align="right">
          <input type="hidden" name="vNext"          value="<%=vNext%>">
          <input type="hidden" name="vEdit"          value="<%=vEdit%>">
          <input type="hidden" name="vCustId"        value="<%=vCustId%>">
          <input type="hidden" name="vCurList"       value="<%=vCurList%>">
          <input type="hidden" name="vGlobal"        value="<%=vGlobal%>">
          <input type="hidden" name="vActive"        value="<%=vActive%>">
          <input type="hidden" name="vFind"          value="<%=vFind%>">
          <input type="hidden" name="vFindFirstName" value="<%=vFindFirstName%>">
          <input type="hidden" name="vFindLastName"  value="<%=vFindLastName%>">
          <input type="hidden" name="vFindEmail"     value="<%=vFindEmail%>">
          <input type="hidden" name="vFindMemo"      value="<%=vFindMemo%>">
          <input type="hidden" name="vFindCriteria"  value="<%=vFindCriteria%>">
          <input type="hidden" name="vLastValue"     value="<%=vMemb_LastName & vMemb_FirstName & vMemb_No%>">
          <input type="hidden" name="vFormat"        value="<%=vFormat%>">
          <input type="hidden" name="vLearners"      value="<%=vLearners%>">
          <input type="submit" name="bNext"          value="<%=bNext%>" class="button085">
        </td>
        <% End If %>

      </tr>
    </table>



    </form>
    <table border="0" width="100%" id="table2" cellspacing="0" cellpadding="10">
      <tr>
        <td align="center">
          <% If vCust_Id = svCustId and vCust_InsertLearners Then %>
          <h2><a href="<%=vEdit%>?vMembNo=0&vNext=<%=vNext%>&vCustId=<%=vCustId%>"><!--webbot bot='PurpleText' PREVIEW='Add a Learner'--><%=fPhra(000370)%></a></h2>
          <% End If %>
          <h2 align="center"><%=vCust_Id & "  (" & vCust_Title & ")"%></h2>
        </td>
      </tr>
    </table>
  </div>
  

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

