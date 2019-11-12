<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/QueryString.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<%
  Dim vNoSource
  If Len(svCustReturnUrl) = 0 Then vNoSource = True Else vNoSource = False
  sGetCust svCustId '...get corporate info
%>


<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %> 

  <div align="center">
    <table border="0" width="60%" cellspacing="0" cellpadding="5">
      <tr>
        <td align="center">
          <h1><!--[[-->Thank you<!--]]-->.</h1>
          <h1><!--[[-->Here is your<!--]]-->&nbsp;<%=fIf(vNoSource, "Customer Id and", "") %>&nbsp;<!--[[-->new Password providing you<!--]]-->&nbsp;<%=vCust_EcomCorpDuration%>&nbsp;<!--[[-->days access to<!--]]-->&nbsp;<%=svCustTitle%>.</h1>

          <div align="center">
            <table border="1" cellpadding="2" cellspacing="0" bordercolor="#FF0000" id="table3">
              <tr>
                <td>

          <table border="0" id="table4">
            <tr>
              <th align="right"><!--[[-->Customer Id<!--]]--> : </th>
              <td><%=Session("EcomCust")%></td>
            </tr>
            <tr>
              <th align="right"><!--[[-->Password<!--]]--> : </th>
              <td><%=Session("EcomId")%></td>
            </tr>
          </table>

                </td>
              </tr>
            </table>
          </div>

          <h2><!--[[-->To begin click <b>Continue</b> where you can enter your<!--]]-->&nbsp;<%=fIf(vNoSource, "Customer Id and", "") %>&nbsp;<!--[[-->new Password.<!--]]--></h2>
          <p><input type="button" onclick="location.href='<%=fIf(vNoSource, "//" & svHost, svCustReturnUrl) %>'" value="Continue" name="B2" class="button"></p>

          <h2 align="center"><!--[[-->You can also click below to automatically <b>Sign In</b>.<!--]]--></h2>
          <p align="center"><input onclick="location.href='//<%=svHost%>/default.asp?vCust=<%=Session("EcomCust")%>&vId=<%=Session("EcomId")%>'" type="button" value="Sign In" name="B1" class="button"></p>

          <table border="0" style="border-collapse: collapse" cellpadding="5" id="table2">
            <tr>
              <td bgcolor="#DDEEF9" valign="top"><a <%=fstatx%> href="javascript:window.print();"><img border="0" src="../Images/Icons/Printer.gif"></a></td>
              <td bgcolor="#DDEEF9"><!--[[-->Print this page for your records.<!--]]--></td>
            </tr>
            <tr>
              <td bgcolor="#DDEEF9" valign="top"><img border="0" src="../Images/Icons/Bang.gif"></td>
              <td bgcolor="#DDEEF9"><!--[[-->Remember to <b>Sign Off</b> after every session.<!--]]--></td>
            </tr>
            <tr>
              <td bgcolor="#DDEEF9" valign="top"><a <%=fstatx%> href="mailto:<%= fIf(Len(svCustEmail) > 0, svCustEmail, "support@vubiz.com")%>?subject=Ecommerce Issue"><img border="0" src="../Images/Icons/Email3.gif"></a></td>
              <td bgcolor="#DDEEF9"><!--[[-->Feel free to email us if you have any questions.<!--]]--></td>
            </tr>
          </table>

        </td>
      </tr>
      </table>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>