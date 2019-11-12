<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vMembId, vMessage : vMessage = ""
  
  '...used in translation engine to Id Type
  p0 = fIf(svCustPwd, fPhraH(000411), fPhraH(000211))

  If Request.Form("vHidden").Count = 1 Then
    sExtractMemb
    vMembId = Request("vMembId")
    If spMembExistsById (svCustAcctId, vMemb_Id) And (vMemb_No = 0 Or vMembId <> vMemb_Id) Then       
      vMessage = fPhraH(001213)
      If vMemb_Id <> vMembId Then vMemb_Id = vMembId '...put back original ID
    Else
      sAddMemb  svCustAcctId

    End If
  End If
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  is th<% Server.Execute vShellHi %>
  <form method="POST" action="UserDemo.asp" target="_self">

    <input type="hidden" name="vHidden"         value="Y">
    <input type="hidden" name="vMemb_No"        value="0">
    <input type="hidden" name="vMembId"         value="<%=vMemb_Id%>">
    <input type="hidden" name="vMemb_FirstName" value="Sample">
    <input type="hidden" name="vMemb_LastName"  value="Certificate">

    <table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td width="100%" valign="top" colspan="2" align="center">
        <h1>Add a Temporary Learner</h1>
        <h2>This allows you to add a temporary guest learner which has all the rights of a normal learner <br>except that any certificates issued show &quot;Sample Certificate&quot; since no Names are assigned to this ID.&nbsp; <br>Note you cannot If you enter a <%=fIf(svCustPwd, "Learner Id", "Password")%> that is on file then any Program.
        <span class="c5"><%=fIf(Len(vMessage)>0, "<br><br>" & vMessage & "<br>" , "")%></span>
        </h2>
        </td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top"><%=fIf(svCustPwd, "Learner Id", "Password")%> : </th>
        <td width="75%" valign="top"><input type="text" size="32" name="vMemb_Id" maxlength="64"> <br>Must be unique - typically use email address.</td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top">Programs : </th>
        <td width="75%" valign="top"><input type="text" size="72" name="vMemb_Programs" maxlength="8000"><br>Enter valid Program Ids separated by spaces, ie: P0011EN P0012EN.</td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top">Programs Expire : </th>
        <td width="75%" valign="top"><input type="text" name="vMemb_Expires" size="12"> <br>Defaults to 7 days from today (<%=fFormatSqlDate(fDefault(vMemb_Expires, Now + 7))%>).</td>
      </tr>
      <tr>
        <td align="center" width="100%" valign="top" colspan="2">
        <p><br><br>When you click Update you can assume this ID was entered and you can continue to add another ID.&nbsp; <br>Click on the Learner Report below to see this learner in the overall Learner Table.<br><br> 
          <input border="0" src="../Images/Buttons/Update_<%=svLang%>.gif" name="vUpdate" type="image">
          <br><br><a href="Users_o.asp?vFind=S&vFindId=<%=vMemb_Id%>&vNext=UserDemo.asp">Learner Report</a><br><br></p>
        </td>
      </tr>
    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


