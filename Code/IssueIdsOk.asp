<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vNoIds, vPrograms, vDuration, vDateFrom, vMembExpires, vMembDuration

  vNoIds    = Request.Form("vNoIds")

  vPrograms = Ucase(Request.Form("vPrograms"))
  vDuration = Request.Form("vDuration")
  vDateFrom = Request.Form("vDateFrom")

  If vDateFrom = "Today" Then
    vMembExpires  = DateAdd("d", vDuration, Now)
    vMembDuration = 0
  ElseIf vDateFrom = "Access" Then
    vMembExpires  = Null
    vMembDuration = vDuration
  Else
    vMembExpires  = Null
    vMembDuration = 0
  End If

%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <base target="_self">
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <form method="POST" action="IssueIdsList.asp">
    <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td colspan="4"><h1 align="left">Generate Multiple Access Ids</h1><h2 align="left">You may enter each learner&#39;s name beside their Password and click update below, or you can add/edit a learner&#39;s name later via &quot;Learner List&quot; report.&nbsp; Please print out this page for your reference. A full listing of all Issued Ids is available using the &quot;Learner List&quot; report. </h2><h6 align="center">Do not refresh this page or new Passwords will be issued in error !</h6></td>
      </tr>
      <tr>
        <th width="20%" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF" nowrap align="left"><%=fIf(svCustPwd, fPhraH(000411), fPhraH(000211))%></th>
        <th width="20%" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF" nowrap align="left">First Name</th>
        <th width="20%" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF" nowrap align="left">Last Name</th>
        <th width="40%" bgcolor="#DDEEF9" height="30" bordercolor="#FFFFFF" nowrap align="left">Email Address</th>
      </tr>
      <%
        vMemb_Programs    = vPrograms
        vMemb_Expires     = vMembExpires
        vMemb_Duration    = vMembDuration
        vMemb_FirstVisit  = ""

        '...Determine the expirey date for the access record
        For i = 1 to vNoIds

          vMemb_No = 0
          vMemb_Id = ""
          sAddMemb svCustAcctId

      %>
      <tr>
        <input type="hidden" size="20" name="vMemb-<%=vMemb_No%>" value="<%=vMemb_Id%>"></td>
        <td width="20%"><%=vMemb_Id%>&nbsp; </td>
        <td width="20%"><input type="text" size="20" name="vFrst-<%=vMemb_No%>"></td>
        <td width="20%"><input type="text" size="20" name="vLast-<%=vMemb_No%>"></td>
        <td width="40%"><input type="text" size="24" name="vEmai-<%=vMemb_No%>"></td>
      </tr>
      <%
        Next
     %>
      <tr>
        <td colspan="4" align="center"><br><input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="I1" type="image"><br>&nbsp;</td>
      </tr>
    </table>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


