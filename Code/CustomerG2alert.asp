<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<%
  Dim vMsg
  sGetCust svCustId
  If Request("vCust_EcomG2alert").Count = 1 Then
    vCust_EcomG2alert = Request("vCust_EcomG2alert")
    sUpdateCustG2alert svCustAcctId, vCust_EcomG2alert
    vMsg = "Updated Successfully"
  End If
%>


<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

  <% 
    Server.Execute vShellHi
  %>
  <table border="0" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="10">
    <form method="POST" action="CustomerG2Alert.asp" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1">
      <tr>
        <td align="center" valign="Top" width="100%" colspan="2">
        <h1 align="center">Customer Email Alert</h1>
        <h2 align="center">This allows you to turn on or off the automatic Email Alerts <br>when content is added to a Learner&#39;s Profile in a Group 2 Ecommerce site.</h2>
        <% =fIf (Len(vMsg) > 0, "<h5 align='center'>" & vMsg & "</h5>", "") %></td>
      </tr>
      <tr>
        <th align="right" width="50%">
        <!--webbot bot='PurpleText' PREVIEW='Include Email Alerts'--><%=fPhra(000934)%> : </th>
        <td width="50%">
          <input class="c2" type="radio" value="1" <%=fCheck(fSqlBoolean(vCust_EcomG2alert), 1)%> name="vCust_EcomG2alert"><!--webbot bot='PurpleText' PREVIEW='Yes'--><%=fPhra(000024)%>&nbsp;&nbsp; 
          <input class="c2" type="radio" value="0" <%=fCheck(fSqlBoolean(vCust_EcomG2alert), 0)%> name="vCust_EcomG2alert"><!--webbot bot='PurpleText' PREVIEW='No'--><%=fPhra(000189)%>&nbsp; </td>
      </tr>
      <tr>
        <th width="100%" colspan="2" height="60"><input type="submit" value="Update" name="bUpdate" class="button"></th>
      </tr>
    </form>
  </table>
  <% 
    Server.Execute vShellLo
  %>

</body>

</html>


