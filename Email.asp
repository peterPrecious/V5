<!--#include virtual = "V5/MailServer/MailServer.asp"-->
<% 
  Dim svMailServer
  sMailServer

  '...this routine allows modules to send emails, sample...
  '   //localhost/v5/email.asp?vFromFirstName=Peter&vFromLastName=Bulloch&vBodyText=Hello Vous ';;LKÉ.;Èpè;lasdf a asdè;lasdè;lfpàçàçàçè È&vFromEmailAddress=pbulloch@vubiz.com&vSubject=Notification&vTrack=Y&vToFirstName=Helen&vToLastName=Eggleston&vToEmailAddress=pbulloch@vubiz.com

  Dim oMail
  Set oMail = Server.CreateObject("SMTPsvg.Mailer")  
  oMail.ClearRecipients      
  oMail.ClearBodyText
  oMail.ContentType  = "text/html"
  oMail.BodyText     = fFr(Request("vBodyText"))
  oMail.FromName     = fFr(Trim(Request("vFromFirstName") & " " & Request("vFromLastName")))
  oMail.FromAddress  = fFr(Request("vFromEmailAddress"))
  oMail.Subject      = fFr(Request("vSubject"))
  oMail.RemoteHost   = svMailServer
  oMail.AddRecipient fFr(Trim(Request("vToFirstName") & " " & Request("vToLastName"))), Request("vToEmailAddress")
  
  If oMail.SendMail Then 
    If Ucase(Request("vTrack")) = "Y" Then Response.Write "OK"
  Else
    If Ucase(Request("vTrack")) = "Y" Then Response.Write oMail.Response
  End If
  
  Set oMail = Nothing

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