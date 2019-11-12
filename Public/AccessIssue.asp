<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Fathmail.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vCust, vEmail, vId, vSubject, vBody, vSender, vRecipients, vStatus 
  If Request.Form.Count > 0 Then
    vCust             = Request("vCust")
    vEmail            = Request("vEmail")   
    vId               = fMembIdByEmail(Right(vCust, 4), vEmail)

    If vId = "AcctId"   Then Response.Redirect "/V5/Code/Error.asp?vErr=" & Server.UrlEncode("That Customer Id (Account) is not valid.")
    If vId = "None"     Then Response.Redirect "/V5/Code/Error.asp?vErr=" & Server.UrlEncode("You are not registered with that Email Address at the Account.")
    If vId = "Multiple" Then Response.Redirect "/V5/Code/Error.asp?vErr=" & Server.UrlEncode("That Email Address is not unique.")

    vSubject        = "Your Vubiz Password"
    vBody           = "<br><br>Your Learner Id (Password) is '" & vId & "'.<br><br><br>" & vbCrLf
    vSender         = "Vubiz Service <info@vubiz.com>"
    vRecipients     = "Vubiz Learner <" & vEmail & ">"
    vStatus         = fFathMail(vSubject, vBody, vSender, vRecipients) 

    Response.Redirect "/V5/Code/Error.asp?vErr=" & Server.UrlEncode("Your Learner Id (Password) has been emailed to the Email Address entered.")

  End If  
%> 


<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="//vubiz.com/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title></title>
  <script>
    function Validate(theForm) {
      if (theForm.vCust.value == ""){
        var vMsg = "Please enter your Customer (Account) ID.";
        alert(vMsg);
        theForm.vCust.focus();
        return (false);
      }
      if (theForm.vEmail.value == ""){
        var vMsg = "Please enter your Email Address.";
        alert(vMsg);
        theForm.vEmail.focus();
        return (false);
      }
      return (true);
    }
  </script>
  <base target="_self">
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <table cellpadding="10" width="100%" border="0" cellspacing="0" id="table11">
    <tr>
      <td width="100%" align="center"><h1><br>Forget Your Password?</h1><h2>Please enter your Customer Id and Email Address then click <b>Go</b>.<br>Your password will be emailed to this address below.<br>&nbsp;</h2>
      <table cellpadding="6" border="0" id="table12" style="border-collapse: collapse" bordercolor="#DDEEF9">
        <form method="POST" action="AccessIssue.asp"  onsubmit="return Validate(this)" name="fForm">
          <tr>
            <td class="c2" nowrap align="right">Customer Id:</td>
            <td class="c2"><input type="text" name="vCust" size="14" value="<%=Request("vCust")%>" maxlength="8" class="c2"></td>
          </tr>
          <tr>
            <td class="c2" nowrap align="right">Email Address:</td>
            <td class="c2"><input type="text" name="vEmail" size="23" value="<%=Request("vEmail")%>" maxlength="64" class="c2"></td>
          </tr>
          <tr>
            <td class="c2" nowrap align="right">&nbsp;</td>
            <td class="c2" align="right"><input type="submit" value="GO" name="bGo" class="button"></td>
          </tr>
        </form>
      </table>
      </td>
    </tr>
  </table>

</body>

</html>


