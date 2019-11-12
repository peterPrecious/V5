<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<% 
  Dim vMsg, vBody, oMail, vOnline
  vMsg = ""

  '...set since bypassing "signin"
  Session("HostDb") = "V5_Vubz"

  '...initially get from Url then from Form - if testing, get from session variable
  vCust_Id             = Request("vCust")
  If fNoValue(vCust_Id) Then 
    vCust_Id           = svCustId
  End If
  
  vMemb_Email          = fUnquote(Request.Form("vMemb_Email"))
  vMemb_FirstName      = fUnquote(Request.Form("vMemb_FirstName"))
  vMemb_LastName       = fUnquote(Request.Form("vMemb_LastName"))
  vMemb_Memo           = fUnquote(Request.Form("vMemb_Memo"))

  sGetCust(vCust_Id)  
  If vCust_Eof Then
    Response.Redirect "Error.asp?vReturn=n&vErr=" & Server.UrlEncode("This service is not currently in service, please email us for support.")
  End If
  
  If fNoValue(vCust_IssueIdsTemplate) Then
    vCust_IssueIdsTemplate = "E0000"
  End If

' response.write fGetEmailBody(vCust_IssueIdsTemplate)

  '...fMembRegister will return a vaid id ("mid")
  If Request.Form("vOnline").Count > 0 Then 

    '...ensure info is valid, if not vMsg will be filled
    vMemb_Id = fMembRegister (vCust_Id)
    
    If Len(vMemb_Id) > 0 Then
  
      vBody = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"

      '...get the template defined in the customer file  
      vBody = vBody & fGetEmailBody(vCust_IssueIdsTemplate)

      vBody = Replace(vBody, "^frs^", vMemb_FirstName)
      vBody = Replace(vBody, "^lst^", vMemb_LastName)
      vBody = Replace(vBody, "^cid^", vCust_Id)
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
        Response.Redirect "Error.asp?vReturn=n&vErr=" & Server.UrlEncode("You are now registered.  Please check your email.  Thank you.")
      Else
        Response.Redirect "Error.asp?vReturn=n&vErr=" & Server.UrlEncode("Sorry, we could not enroll you becuase there was a proble sending the email message! The address may be invalid or the Mail Server may be unavailable.")
      End If
    
    End If
  End If  

  Function fGetEmailBody(vTemplate)
    Const vForReading = 1, vForWriting = 2
    Dim oFs, oF, oFile, vFile
    vFile = Lcase(Server.MapPath("\V5\Features\EmailTemplates\" & vTemplate & ".txt"))
    Set oFs   = CreateObject("Scripting.FileSystemObject")
    Set oF    = oFs.GetFile(vFile)
    Set oFile = oFs.OpenTextFile(vFile, vForReading)
    fGetEmailBody = oFile.ReadLine
  End Function

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
  <form method="POST" action="IssueIdsByEmail.asp">
    <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="0" cellspacing="0">
      <tr>
        <td colspan="2">
        <h1 align="center">Issue a Password by Email</h1>
        <h2>Please enter your name and the email address where you wish us to send your system generated password.&nbsp; Once you click &quot;continue&quot; you will be enrolled.&nbsp; Please allow a few minutes for the email system to deliver you password.</h2>
        </td>
      </tr>
      <% If vMsg <> "" Then %>
      <tr>
        <td colspan="2" align="center"><%=vMsg%>&nbsp; </td>
      </tr>
      <% End If %>
      <tr>
        <th align="right" width="30%">First Name : </th>
        <td width="70%">&nbsp;<input type="text" name="vMemb_FirstName" size="25"></td>
      </tr>
      <tr>
        <th align="right" width="30%">Last Name : </th>
        <td width="70%">&nbsp;<input type="text" name="vMemb_LastName" size="25"></td>
      </tr>
      <tr>
        <th align="right" width="30%">Email Address : </th>
        <td width="70%">&nbsp;<input type="text" name="vMemb_Email" size="45"></td>
      </tr>
      <% If vCust_IssueIdsMemo Then %>
      <tr>
        <th align="right" width="30%">Memo : </th>
        <td width="70%">&nbsp;<input type="text" name="vMemb_Memo" size="45"></td>
      </tr>
      <% End If %>
      <tr>
        <td align="center" colspan="2"><br>&nbsp;&nbsp;&nbsp;&nbsp; <a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <input border="0" src="../Images/Buttons/Continue_<%=svLang%>.gif" name="I1" type="image"><br>&nbsp;</td>
      </tr>
    </table>
    <input type="hidden" name="vOnline" value="y"><input type="hidden" name="vCust" value="<%=vCust_Id%>">
  </form>

  <% Server.Execute vShellLo %>

</body>

</html>
