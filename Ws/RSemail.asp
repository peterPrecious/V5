<%
Option Explicit
If Request.Form.Count <> 0 then
  Dim c, i, vBody, vError, oEmail, vSub, vTemp, vButMsg
  '... Set true for further processing
  vError = False
  vSub = True
  c = request.form.count

  For i = 1 to c
    '... Be sure to omit sending the Submit button or Image info.
    If NOT UCase(Request.Form.Key(i)) = "SUBMIT" AND NOT UCase(Request.Form.Key(i)) = "FROM" AND NOT UCase(Request.Form.Key(i)) = "RECIPIENT" AND NOT UCase(Request.Form.Key(i)) = "SUBJECT" then
      If Instr(Request.Form.Key(i), "<P>") then
        vTemp = "<BR>" & Replace(Mid(Request.Form.Key(i), 4), "_", " ") & Request.Form(i) & "<BR>"
      Else
        vTemp = Replace(Request.Form.Key(i), "_", " ") & " " & Request.Form(i) & "<BR>"
      End If
      vBody = vBody & vTemp
    End If
  Next
  
  Set oEmail = Server.CreateObject("SMTPsvg.Mailer")
  oEmail.ContentType    = "text/html"
  oEmail.FromName       = Request.Form("From")
  oEmail.FromAddress    = Request.Form("From")
  oEmail.RemoteHost     = "mail.dades.ca; windex.dades.ca"
  oEmail.ReturnReceipt  = false
  oEmail.ConfirmRead    = false
  oEmail.Subject        = Request.Form("Subject")
  oEmail.ClearBodyText
  oEmail.BodyText       = vBody
  oEmail.Recipient      = "koh@vubiz.com"
  If Not oEmail.SendMail Then vError = True
  oEmail.ClearRecipients
  oEmail.ClearBodyText    
  Set oEmail            = Nothing 

End If
%>

<html>

<head>
  <title>Thank you</title>
</head>

<body>

  <p align="center">&nbsp;</p><p align="center">&nbsp;</p><p align="center">&nbsp;</p><%If Not vError then%> <p align="center"><font face="Arial" color="#800000"><b>Thank you, your submission has been received.</b></font></p><%Else%> <p align="center"><font face="Arial" color="#800000"><b>An error has occurred, your submission could not be processed.<br>Please email <a href="<%=Request.Form("Recipient")%>"><%=Request.Form("Recipient")%></a><br>Thank You</b></font></p><%End If%>
  <div align="center">
    <center>
    <table border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="100%">
        <form id="form1" name="form1">
          <input type="button" name="Close" onclick="window.close()" value="Close Window">
        </form>
        </td>
      </tr>
    </table>
    </center></div>

</body>

</html>
