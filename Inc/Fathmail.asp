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