<% 
  Sub sDebugByEmail_Old (vSubject, vBody)
    Dim vHtml, oMail
    Set oMail = Server.CreateObject("SMTPsvg.Mailer")  
    vHtml = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"
    vHtml = vHtml & "<html><head><meta http-equiv='Content-Language' content='en-us'><title></title></head><body><p><font face='Verdana' size='1' color='#000080'>"
    vHtml = vHtml & vBody
    vHtml = vHtml & "</font></body></html>"
    oMail.ClearRecipients      
    oMail.ClearBodyText
    oMail.ContentType  = "text/html"
    oMail.BodyText     = vHtml
    oMail.FromName     = "Vubiz QA Test"
    oMail.FromAddress  = "pbulloch@vubiz.com"
    oMail.Subject      = vSubject
    oMail.RemoteHost   = svMailServer
'   oMail.AddRecipient "Helen Eggleston", "heggleston@vubiz.com"
'   oMail.AddRecipient "Allison Lee", "alee@vubiz.com"
'   oMail.AddRecipient "Peter Bulloch", "pbulloch@vubiz.com"
    oMail.SendMail     
  End Sub


  Sub sDebugByEmail (vSubject, vBody)

    Dim oXmlHttp, vXmlUrl, vXmlCommand, vXmlResponse
    Dim vToFirstName, vToLastName, vFromEmailAddress, vFromFirstName, vFromLastName, vToEmailAddress, vBodyText
  

    '...select live or test block of vTo...

'   vToFirstName      = "Peter"
'   vToLastName       = "Bulloch"
'   vToEmailAddress   = "pbulloch@vubiz.com"
'   vToEmailAddress   = "peter@bullochonline.com"

    vToFirstName      = "Allison"
    vToLastName       = "Lee"
    vToEmailAddress   = "alee@vubiz.com"

    vFromEmailAddress = "salessupport@vubiz.com"
    vFromFirstName    = Server.UrlEncode("Vubiz Ecom Scripts")
    vFromLastName     = ""

    vBodyText         = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>"
    vBodyText         = vBodyText & "<html><head><meta http-equiv='Content-Language' content='en-us'><title></title></head><body><p><font face='Verdana' size='1' color='#000080'>"
    vBodyText         = vBodyText & vBody
    vBodyText         = vBodyText & "</font></body></html>"
    
    vBodyText         = Server.UrlEncode(vBodyText)
  
    vXmlCommand       = "vToFirstName=" & vToFirstName & "&vToLastName=" & vToFirstName & "&vFromEmailAddress=" & vFromEmailAddress & "&vFromFirstName=" & vFromFirstName & "&vFromLastName=" & vFromLastName & "&vToEmailAddress=" & vToEmailAddress & "&vSubject=" & vSubject & "&vBodyText=" & vBodyText & "&vTrack=X&vContentType=text/html"

    vXmlUrl           = "//66.135.100.74/Email/Email.asp"
  
    Set oXmlHttp      = Server.Createobject("MSXML2.ServerXMLHTTP")
    oXmlHttp.Open "POST", vXmlUrl, false
    oXmlHttp.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
    oXmlHttp.Send vXmlCommand
    vXmlResponse = oXmlHttp.ResponseText
    Set oXmlHttp = Nothing
  
    If vXmlResponse <> "OK" Then 
      Response.Write "You email could not be sent : " & vXmlResponse
    End If

  End Sub



















%>