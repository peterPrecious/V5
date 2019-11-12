<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <link href="<%=svDomain%>/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>:: Site Feedback</title>
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080" bgcolor="#ffffff">

  <table border="1" width="100%" cellpadding="10" bordercolor="#DDEEF9" style="border-collapse: collapse" cellspacing="10" bgcolor="#F2F9FD">
    <tr>
      <td align="center">
      <h1>Site Feedback<br>&nbsp;</h1>
      <h2 align="left">Please feel free to send us your comments about this page or any elements of this site or any problems you may be be experiencing.&nbsp; Your feedback is appreciated!&nbsp;
      <% ="(" & svMembFirstName & " " & svMembLastName & fIf(Len(svMembEmail)> 0 , " - " & svMembEmail, "") & ")" %></h2>
      <h2>Please enter your comments then click <b>Send</b>.</h2>
      <p>
      <textarea rows="8" name="vSite_Feedback" cols="60">
      </textarea></p>
      <p>
      <br>
      <input type="button" onclick="alert('Not yet operational... \ncoming in the fullness of time, \nat the appropriate juncture.')" value="Send" name="bSend" class="button"><h2>
      Click the X at the top right if you do not wish to forward any comments.</h2>
      </td>
    </tr>
  </table>
    
</body>

</html>