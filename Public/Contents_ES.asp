<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<html>

<head>
  <title>vuNews</title>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <base target="Mkt_Main">
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">


  <div align="center">
    <br>&nbsp;<table cellspacing="0" cellpadding="6" border="1" id="table1" style="border-collapse: collapse" bordercolor="#DDEEF9" width="150">
      <form method="GET" action="../Default.asp" target="_top">
        <tr>
          <td class="c2">
            <p class="c2">Customer Id:<br>
            <input type="text" name="vCust" size="14"><br>Password:<br>
            <input type="password" name="vId" size="14"></p><p align="right">
            <input type="submit" value="Sign In" name="bGo" class="button">
            <input type="hidden" name="vLang" value="EN">
          </td>
        </tr>
      </form>
    </table>
    </div>

  <div align="center">
    <table border="0" cellspacing="0" cellpadding="5" bordercolor="#DDEEF9" class="c2" width="150">
      <tr>
        <td class="c2"><br><b><a href="Archives/01_Welcome_EN.asp">Welcome</a></b></td>
      </tr>
      <tr>
        <td class="c2"><b><a target="_top" href="Cat_Start.asp">Catalogue</a></b></td>
      </tr>

<!--
      <tr>
        <td><a href="Subscriber.asp?vAction=add">Subscribe to VuNews</a></td>
      </tr>
      <tr>
        <td><a href="Subscriber.asp?vAction=edit">Update My Profile</a></td>
      </tr>
      <tr>
        <td><a href="Subscriber.asp?vAction=del">Unsubscribe to VuNews</a></td>
      </tr>
      <% If Session("vuNewsAdmin") = "Ok" Then %>
      <tr>
        <td><a href="SubscriberReport.asp">Subscriber Report</a></td>
      </tr>
      <tr>
        <td><a href="SubscriberEmail.asp">Email Subscribers</a></td>
      </tr>
      <% End If%>
      <tr>
        <td align="center"><a href="SubscriberSignIn.asp">&#8486;</a></td>
      </tr>
  


-->

    </table>
  </div>


  <br>

  </body>

</html>
