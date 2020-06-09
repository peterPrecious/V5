<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<% Response.Redirect "/Chaccess/Signin" %>

<%
  '...if house account do not display, it will be (re)assigned at signin
  Dim vCust
  vCust = Request("vCust")
  If vCust = "VUBZ2274" Then vCust = ""  
%>  

<html>

<head>
  <title>:: Vubiz</title>
  <meta charset="UTF-8">
  <link href="http://vubiz.com/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <script>
    function validate(theForm) {
      if (theForm.vId.value == ""){
        var vMsg = "Please enter a valid Password.";
        alert(vMsg);
        theForm.vId.focus();
        return (false);
      }
      return (true);
    }


    function browserCheck(url) {
      var modwindow = window.open(url,'Module','toolbar=no,width=780,height=546,left=10,top=10,status=yes,scrollbars=yes,resizable=yes')
    }  

  </script>
  <base target="_self">
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <div align="center">
      <table cellpadding="2" border="0" id="table12" style="border-collapse: collapse" bordercolor="#DDEEF9">

        <form method="GET" action="../Default.asp" target="_top" onsubmit="return validate(this)" name="fForm">
          <input type="hidden" name="vLang" value="EN">
          <tr>
            <th class="c2" colspan="2">
              &nbsp;<h1>Login</h1><h2>Please enter your Customer Id and Password,<br>then click <b>Login</b>.</h2>
            </th>
          </tr>
          <tr>
            <th class="c2" align="right">
              Customer Id :</th>
            <td class="c2">
                <input type="text" name="vCust" size="10" value="<%=Request("vCust")%>" class="c2" maxlength="8"></td>
          </tr>
          <tr>
            <th class="c2" align="right">
              Password :</th>
            <td class="c2">
                <input type="password" name="vId" size="23" value="<%=Request("vId")%>" maxlength="64" class="c2"></td>
          </tr>
          <tr>
            <th class="c2" align="right">
              &nbsp;</th>
            <td class="c2" align="right">
                <input type="submit" value="Login" name="bGo" class="button70"></td>
          </tr>
          <tr>
            <td class="c2" colspan="2">
              <br>Helpful links....
              <ul>
                <li><a target="_top" onclick="browserCheck('http://www.kmsi.us/bh/dotnet/vb/kmxdirect.aspx?bhcp=1')" href="/v5/Public/01_Welcome_EN.asp?vFrame=10_Login_EN.asp">Browser Check</a><br>&nbsp;</li>
                <li><a target="_self" href="AccessIssue.asp">Lost Password</a><br>&nbsp;</li>
                <li><a href="BrowserIssues.htm">Help Accessing this Service</a><br>&nbsp;</li>
              </ul>
            </td>
          </tr>
        </form>

      </table>
      
</body>

</html>
