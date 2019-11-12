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
          <h1><!--webbot bot='PurpleText' PREVIEW='Thank you'--><%=fPhra(000246)%>.</h1>
          <h1><!--webbot bot='PurpleText' PREVIEW='Here is your'--><%=fPhra(000348)%>&nbsp;<%=fIf(vNoSource, "Customer Id and", "") %>&nbsp;<!--webbot bot='PurpleText' PREVIEW='new Password providing you'--><%=fPhra(000358)%>&nbsp;<%=vCust_EcomCorpDuration%>&nbsp;<!--webbot bot='PurpleText' PREVIEW='days access to'--><%=fPhra(000359)%>&nbsp;<%=svCustTitle%>.</h1>

          <div align="center">
            <table border="1" cellpadding="2" cellspacing="0" bordercolor="#FF0000" id="table3">
              <tr>
                <td>

          <table border="0" id="table4">
            <tr>
              <th align="right"><!--webbot bot='PurpleText' PREVIEW='Customer Id'--><%=fPhra(000111)%> : </th>
              <td><%=Session("EcomCust")%></td>
            </tr>
            <tr>
              <th align="right"><!--webbot bot='PurpleText' PREVIEW='Password'--><%=fPhra(000211)%> : </th>
              <td><%=Session("EcomId")%></td>
            </tr>
          </table>

                </td>
              </tr>
            </table>
          </div>

          <h2><!--webbot bot='PurpleText' PREVIEW='To begin click <b>Continue</b> where you can enter your'--><%=fPhra(000349)%>&nbsp;<%=fIf(vNoSource, "Customer Id and", "") %>&nbsp;<!--webbot bot='PurpleText' PREVIEW='new Password.'--><%=fPhra(000350)%></h2>
          <p><input type="button" onclick="location.href='<%=fIf(vNoSource, "//" & svHost, svCustReturnUrl) %>'" value="Continue" name="B2" class="button"></p>

          <h2 align="center"><!--webbot bot='PurpleText' PREVIEW='You can also click below to automatically <b>Sign In</b>.'--><%=fPhra(000345)%></h2>
          <p align="center"><input onclick="location.href='//<%=svHost%>/default.asp?vCust=<%=Session("EcomCust")%>&vId=<%=Session("EcomId")%>'" type="button" value="Sign In" name="B1" class="button"></p>

          <table border="0" style="border-collapse: collapse" cellpadding="5" id="table2">
            <tr>
              <td bgcolor="#DDEEF9" valign="top"><a <%=fstatx%> href="javascript:window.print();"><img border="0" src="../Images/Icons/Printer.gif"></a></td>
              <td bgcolor="#DDEEF9"><!--webbot bot='PurpleText' PREVIEW='Print this page for your records.'--><%=fPhra(000346)%></td>
            </tr>
            <tr>
              <td bgcolor="#DDEEF9" valign="top"><img border="0" src="../Images/Icons/Bang.gif"></td>
              <td bgcolor="#DDEEF9"><!--webbot bot='PurpleText' PREVIEW='Remember to <b>Sign Off</b> after every session.'--><%=fPhra(000347)%></td>
            </tr>
            <tr>
              <td bgcolor="#DDEEF9" valign="top"><a <%=fstatx%> href="mailto:<%= fIf(Len(svCustEmail) > 0, svCustEmail, "support@vubiz.com")%>?subject=Ecommerce Issue"><img border="0" src="../Images/Icons/Email3.gif"></a></td>
              <td bgcolor="#DDEEF9"><!--webbot bot='PurpleText' PREVIEW='Feel free to email us if you have any questions.'--><%=fPhra(000276)%></td>
            </tr>
          </table>

        </td>
      </tr>
      </table>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

