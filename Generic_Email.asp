<!--#include virtual = "V5\MailServer\MailServer.asp"-->
<% 
  Dim svMailServer
  sMailServer
 
  If Request.Form.Count <> 0 Then

    Dim oMail, vHost, vMailServer
    Dim c, i, vBody, vError, vSub, vTemp, vButMsg
 
    '... Set true for further processing
    vError = False
    vSub = True
    c = request.form.count
  
    For i = 1 To c
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
    
    Set oMail = Server.CreateObject("SMTPsvg.Mailer")
    oMail.ContentType    = "text/html"
    oMail.FromName       = fFr(Request.Form("From"))
    oMail.FromAddress    = fFr(Request.Form("From"))
    oMail.RemoteHost     = svMailServer
    oMail.ReturnReceipt  = false
    oMail.ConfirmRead    = false
    oMail.Subject        = fFr(Request.Form("Subject"))
    oMail.ClearBodyText
    oMail.BodyText       = fFr(vBody)
    oMail.Recipient      = fFr(Request.Form("Recipient"))
    If Not oMail.SendMail Then vError = True
    oMail.ClearRecipients
    oMail.ClearBodyText    
    Set oMail            = Nothing 
  
  End If


  Function fFr (vPhrase)
    fFr = vPhrase
    fFr = Replace(fFr, "à", "&#224;") 
    fFr = Replace(fFr, "ç", "&#231;") 
    fFr = Replace(fFr, "è", "&#232;") 
    fFr = Replace(fFr, "é", "&#233;") 
    fFr = Replace(fFr, "ê", "&#234;") 

    fFr = Replace(fFr, "À", "&#192;") 
    fFr = Replace(fFr, "Ç", "&#199;") 
    fFr = Replace(fFr, "È", "&#200;") 
    fFr = Replace(fFr, "É", "&#201;") 
    fFr = Replace(fFr, "Ê", "&#202;") 
  End Function



%>

<html>

<head>
  <title>:: Course Evaluation</title>
</head>

<body>

  <p align="center">&nbsp;</p><p align="center">&nbsp;</p><p align="center">&nbsp;</p><%
    If uCase(Request.QueryString("vLang")) = "FR" then
      vButMsg = "Fermez la fenêtre"
  %> <%  
      If Not vError Then
  %> <p align="center"><font face="Arial" color="#800000"><b>Merci de votre collaboration. Vos commentaires nous sont précieux.</b></font></p><%Else%> <p align="center"><font face="Arial" color="#800000"><b>Une erreur nous empêche de traiter votre sondage. Veuillez avertir VUBIZ.<br>Please email <a href="<%=Request.Form("Recipient")%>"><%=Request.Form("Recipient")%></a><br>Merci</b></font></p><%  
      End If 
  
    Else

      vButMsg = "Close Window"
      
      If Not vError Then
  %> <p align="center"><font face="Arial" color="#800000"><b>Thank you, your submission has been received.</b></font></p><%
      Else
  %> <p align="center"><font face="Arial" color="#800000"><b>An error has occurred, your submission could not be processed.<br>Please email <a href="<%=Request.Form("Recipient")%>"><%=Request.Form("Recipient")%></a><br>Thank You</b></font></p><%
  
     End If
     
   End If
 %>
  <div align="center"><center>
    <table border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="100%">
        <form id="form1" name="form1">
          <input type="button" name="Close" onclick="window.close()" value=" <%=vButMsg%> ">
        </form>
        </td>
      </tr>
    </table>
    </center></div>

</body>

</html>
