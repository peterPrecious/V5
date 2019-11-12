<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity= True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<% 
  '...set since bypassing "signin"
  Session("HostDb") = "V5_Vubz"

  Dim vCust, vMsg, vBody, oMail, vOnline
  vMsg = ""

  '...originally received from some landing page
  vCust                = Request("vCust")
  If Len(vCust) <> 8 Then
    vMsg = "This service is not currently in service, please email us for support."
  End If

  '...from learners form
  vMemb_Email          = Request.Form("vMemb_Email")
  vMemb_FirstName      = Request.Form("vMemb_FirstName")
  vMemb_LastName       = Request.Form("vMemb_LastName")

  '...fMembRegister will return a vaid id ("mid")
  If Request.Form("vOnline").Count > 0 Then 
    '...ensure info is valid, if not vMsg will be filled
    vMemb_Id = fMembRegister (vCust)
    If Len(vMemb_Id) > 0 Then
  
      vBody = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"
      vBody = vBody & "<html><head><meta http-equiv="Content-Type" content="text/html; charset=windows-1252"><meta http-equiv='Content-Language' content='en-us'>  <meta http-equiv="Cache-Control" content="no-cache">
  <meta http-equiv="Pragma" content="no-cache">
  <meta http-equiv="Expires" content="-1"><base target='_self'></head><body leftmargin='0' topmargin='0' bgcolor='#FFFFFF' text='#000080' link='#000080' vlink='#000080' alink='#000080'><div align='center'>  <center>  <table border='0' width='97%' cellspacing='0' cellpadding='0'>    <tr>      <td width='100%' align='right' colspan='3' valign='bottom'><img border='0' src='/V5/Images/Shell/1x1TransparentSpacer.gif' width='50' height='15'></td>    </tr>    <tr>      <td width='100%' align='right' colspan='3' valign='bottom' background='/V5/Images/Shell/HolderTop_Bg.gif'><img border='0' src='/V5/Images/Shell/HolderTop_Right.gif' width='114' height='25'></td>    </tr>    <tr>      <td width='1%' valign='bottom' background='/V5/Images/Shell/HolderLeft_Spacer.gif'><img border='0' src='/V5/Images/Shell/HolderLeft_Spacer.gif' width='25' height='54'></td>      <td width='98%' align='center' valign='middle'>      <table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='100%' id='AutoNumber1'>        <tr>          <td width='100%' align='center'><img border='0' src='/V5/Images/VuBizLogo.gif' width='200' height='76'></td>        </tr>        <tr>          <td width='100%' align='center'>          <font face='Verdana' size='1'><br><br><br>Thank you ^frs^.&nbsp; You are now registered.<br><br>You can access your content at: <a href='//^url^'>//^url^</a>, using: </font><br>&nbsp;          <table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='100%'>            <tr>              <td width='50%' align='right'><font face='Verdana' size='1'>Id&nbsp;:&nbsp;&nbsp;</font> </td>              <td width='50%'><font face='Courier New' size='2'>&nbsp;^cid^</font></td>            </tr>            <tr>              <td width='50%' align='right'><font face='Verdana' size='1'>Password&nbsp;:&nbsp;&nbsp;</font></td>              <td width='50%'><font face='Courier New' size='2'>&nbsp;^mid^</font></td>            </tr>          </table>          </td>        </tr>      </table>      <p>&nbsp;</p>      </td>      <td width='1%' valign='bottom' background='/V5/Images/Shell/HolderRight_Spacer.gif'><img border='0' src='/V5/Images/Shell/HolderRight_Spacer.gif' width='25' height='54'></td>    </tr>    <tr>      <td width='100%' align='right' colspan='3' valign='top' background='/V5/Images/Shell/HolderBottom_Bg.gif'><img border='0' src='/V5/Images/Shell/HolderBottom_Right.gif' width='114' height='25'></td>    </tr>  </table>  </center></div><p>&nbsp;</p></body></html>"

      vBody = Replace(vBody, "^frs^", vMemb_FirstName)
      vBody = Replace(vBody, "^lst^", vMemb_LastName)
      vBody = Replace(vBody, "^cid^", vCust)
      vBody = Replace(vBody, "^mid^", vMemb_Id)
      vBody = Replace(vBody, "^url^", svHost)
  
      Set oMail = Server.CreateObject("SMTPsvg.Mailer")  
      oMail.ClearRecipients      
      oMail.ClearBodyText
      oMail.ContentType  = "text/html" '...note mandatory "html" format
      oMail.BodyText     = vBody
      oMail.FromName     = "Vubiz Registration Systems"
      oMail.FromAddress  = "info@vubiz.com"
      oMail.Subject      = "Your password to " & " Vubiz"
      oMail.RemoteHost   = svMailServer
      oMail.AddRecipient vMemb_FirstName & " " & vMemb_LastName, vMemb_Email
      If oMail.SendMail Then 
        vMsg = "Email Sent"
      Else
        vMsg = "Email could not be sent! The address may be invalid or the Mail Server may be unavailable."
      End If
    
    End If
  End If  
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <base target="_self">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <!-- Display form -->
  
  <% If vMsg = "" Then %>
  <form method="POST" action="IssueIdsByPubicEmail.asp">
    <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="2" align="center">
        <h1>Register by Email</h1>
        <h2>Please enter your name and the email address where you wish us to send your password, then click &quot;continue&quot;.</h2>
        </td>
      </tr>
      <tr>
        <th align="right" width="30%">First Name : </th>
        <td width="70%">&nbsp;<input type="text" name="vMemb_FirstName" size="32"></td>
      </tr>
      <tr>
        <th align="right" width="30%">Last Name : </th>
        <td width="70%">&nbsp;<input type="text" name="vMemb_LastName" size="32"></td>
      </tr>
      <tr>
        <th align="right" width="30%">Email Address : </th>
        <td width="70%">&nbsp;<input type="text" name="vMemb_Email" size="45"></td>
      </tr>
      <tr>
        <td align="center" colspan="2"><br>&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Continue_<%=svLang%>.gif" name="I1" type="image"><h6>NOTE: Please allow a few minutes for the email system to deliver your access instructions.</h6></td>
      </tr>
    </table>
    <input type="hidden" name="vOnline" value="y"><input type="hidden" name="vCust" value="<%=vCust%>">
  </form>
  <!-- Display message - good or bad -->
  
  <% Else %> 
  
  <p align="center"><b><font color="#FF0000" face="Verdana" size="2"><%=vMsg%></font></b></p>
  
  <% End If %>


  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
