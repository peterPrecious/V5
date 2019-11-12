<% 
  sGetQueryString
  If Request.Form("fProfile") = "Y" Then
    vMemb_No        = svMembNo
    vMemb_Pwd       = Ucase(Request.Form("vMemb_Pwd"))
    vMemb_FirstName = Request.Form("vMemb_FirstName")
    vMemb_LastName  = Request.Form("vMemb_LastName") 
    vMemb_Email     = Request.Form("vMemb_Email")
    vMemb_VuNews    = fDefault(Request.Form("vMemb_VuNews"), 0)
    sUpdateMemb_Profile 
    Response.Redirect "#MyProfile"
  End If


  '...this will put either support@vubiz.com or the customers email address on the Contact Us link at the bottom
  Function fContactUs
    Dim vEmail, vText
    If Len(svCustEmail) > 0 Then
      vEmail = svCustEmail
    Else
      vEmail = "support@vubiz.com"
    End If
    Select Case svLang
      Case "FR" : vText = "Communiquez avec nous"
      Case "ES" : vText = "P&#243;ngase en contacto con nosotros"
      Case Else : vText = "Contact Us"
    End Select
    fContactUs = "<a href='mailto:" & vEmail & "?subject=" & svCustId & " Issue'>" & vText & " (" & vEmail & ")</a>"
  End Function
%>