<%
  Function fFathMail(vSubject, vTextHTML, vSender, vRecipients)  

    Server.ScriptTimeout = 60 * 10
  
    Dim oMsg, oMail 
    Set oMsg  = Server.CreateObject("FathMail.Message")
    Set oMail = Server.CreateObject("FathMail.SMTP")
    
    oMsg.Subject      = vSubject
    oMsg.TextHTML     = vTextHTML
    oMsg.Sender       = vSender
    oMsg.Recipients   = vRecipients
  
'   oMail.LoginMethod = 2
'   oMail.Username    = "info@vubiz.com"
'   oMail.Password    = "vubizpass"
    oMail.ServerPort  = 25
    oMail.ServerAddr  = "localhost"
    oMail.Send oMsg
  
    If Err.Number     = 0 Then
    	fFathMail = "Ok"
    Else
    	fFathMail = "Email Error: " & oMail.LastCommandResponse 
    End If
    oMail.Disconnect
  
    Set oMsg  = Nothing
    Set oMail = Nothing
    
  End Function


  Function fFr (vPhrase)
    fFr = vPhrase

    fFr = Replace(fFr, "�", "&#224;") 
    fFr = Replace(fFr, "�", "&#231;") 
    fFr = Replace(fFr, "�", "&#232;") 
    fFr = Replace(fFr, "�", "&#233;") 
    fFr = Replace(fFr, "�", "&#234;") 

    fFr = Replace(fFr, "�", "&#192;") 
    fFr = Replace(fFr, "�", "&#199;") 
    fFr = Replace(fFr, "�", "&#200;") 
    fFr = Replace(fFr, "�", "&#201;") 
    fFr = Replace(fFr, "�", "&#202;") 
  End Function

%>