<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Querystring.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<% 
    sGetQueryString
    Dim vCertType, vCertLogo1, vCertLogo2, vCertLogos, vCertId, vCertTitle, vCertMark, vCertName, vCertDate, vCertTimeSpent, vOk, vFolder
  
    '...if guest learner change name
    If svMembLevel = 1 Then Session("CertName") = fIf(svLang = "FR", "** Certificat échantillon **", "** Sample Certificate **") 
  
    '...Unless this is a sample, display the print alert ...................... do not use, migrating to new certificate
    If svMembLevel = 99 And Session("CertSample") <> "y" Then    
      Dim vScript
      vScript = ""
      vScript = vScript   & "<script for='window' event='onload'>" & vbCrLf
      If Session("CertType") <> "Completion" Then
        vScript = vScript & "  if (opener!=null)" & vbCrLf
        vScript = vScript & "    if (opener.opener!=null)" & vbCrLf
        vScript = vScript & "      if (opener.opener.opener != null) opener.opener.close()" & vbCrLf
      End If
      If vPlatform = "win" Then
        vScript = vScript & "        alert('" & fPhraH(000219) & "')" & vbCrLf
      Else
        vScript = vScript & "        alert('" & fPhraH(000218) & "')" & vbCrLf
      End If
      vScript = vScript &   "</script>"
      Response.Write vScript 
    End If
  

    '...Get cert info from querystring/form/session variables 


    '...get logo configuration from customer table
    sGetCust (vCust)     
    vCertLogo1   = Session("CertLogo1")
    If fNoValue(vCertLogo1) Then
      vCertLogo1 = fOkValue(vCust_CertLogoVubiz)
    End If
    '...If no vubiz logo...
    If Len(vCertLogo1) = 0 Then
      vCertLogo1 = "vubz.gif"
    End If  
    vCertLogo2   = Session("CertLogo2")
    If fNoValue(vCertLogo2) Then
      vCertLogo2 = fOkValue(vCust_CertLogoCust)
    End If
    '...If no client logo...
    If Len(vCertLogo2) = 0 Then
      vCertLogo2 = svCustBanner
    End If
    vCertLogos   = Session("CertLogos") 
    If fNoValue(vCertLogos) Then
      If vCertLogo1     = "n" And vCertLogo2 = "n" Then 
        vCertLogos      = "None"
      ElseIf vCertLogo1 = "n" Then 
        vCertLogos      = "Cust"
      ElseIf vCertLogo2 = "n" Then 
        vCertLogos      = "Vubiz"
      Else
        vCertLogos      = "Both"
      End If
    End If
  

    vCertType    = Request("vCertType")
    If fNoValue(vCertType) Then
      vCertType  = Session("CertType")
    End If
    If fNoValue(vCertType) Then
      vCertType  = "Test"
    End If
    Session("CertType") = vCertType

  
    vCertId      = Request("vModId")
    If fNoValue(vCertId) Then
      vCertId    = Request("vCertId")
    End If
    If fNoValue(vCertId) Then
      vCertId    = Session("CertId")
    End If
    Session("CertId") = vCertId

  
    vCertTitle   = Request("vCertTitle")
    If Len(vCertTitle) = 0 Then
      vCertTitle = Session("CertTitle")
    End If
    If Len(vCertTitle) = 0 Then
      vCertTitle = fModsTitle(vCertId)
    End If
    Session("CertTitle") = vCertTitle
    
  
    vCertMark    = Request("vMark")
    If fNoValue(vCertMark) Then
      vCertMark  = Session("CertMark")
    End If
    Session("CertMark") = vCertMark

    '...If no name
    vCertName   = Request("vCertName")
    If Len(vCertName) = 0 Then
      vCertName = Session("CertName")
    End If
    If Len(vCertName) = 0 Then
      vCertName = Trim(svMembFirstName & " " & svMembLastName)
    End If
    Session("CertName") = vCertName

    '...If no Cert date
    vCertDate    = Request("vCertDate")
    If Not IsDate(vCertDate) Then
      vCertDate  = Session("CertDate")
    End If
    If Not IsDate(vCertDate) Then
      vCertDate = Now
    End If
    vCertDate = FormatDateTime (vCertDate, vbShortDate)
    Session("CertDate") = vCertDate
  
    '...If custom cert note priority: program assessment/customer cert | program repository/tools | customer repository/tools

    '   this If statement will only be bypass if its sent to the custom cert and its NOT an exam.
    If Request("vTempBypass") <> "Bypass" Then    

      '...from repository or custom certs?
      If Len(Session("CertProg")) > 0 Then
        sGetProg (Session("CertProg"))

        '...this will pick up all values from session variables
        If vProg_CustomCert Then 
          Response.Redirect "/V5/Repository/" & svHostDb & "/" & Session("CertProg") & "/Tools/Certificate.asp"
        ElseIf vCust_CustomCert Then 
          Response.Redirect "/V5/Repository/" & svHostDb & "/" & svCustAcctId & "/Tools/Certificate.asp"
        ElseIf Len(vProg_AssessmentCert) > 0 Then
          vFolder = vProg_AssessmentCert 
        ElseIf Len(vCust_AssessmentCert) > 0  Then 
          vFolder = vCust_AssessmentCert 
        Else 
          vFolder = svLang
        End If
        
        Response.Redirect "/V5/Assessments/CustomCerts/" & vFolder & "/Default.htm?vFirstName=" & svMembFirstName & "&vLastName=" & svMembLastName & "&vLastScore=" & fFormatDate(Session("CertDate")) & "&vMods_Id=" & Session("CertId") & "&vMods_Title=" & Session("CertTitle") & "&vScore=" & vCertMark &  "&logo=" & svCustBanner

      End If

    End If
  





   '........................if it gets here then it's a standard platform cert, use new cert

   Response.Redirect "/V5/Assessments/CustomCerts/" & svLang & "/Default.htm?vFirstName=" & svMembFirstName & "&vLastName=" & svMembLastName & "&vLastScore=" & fFormatDate(Session("CertDate")) & "&vMods_Id=" & Session("CertId") & "&vMods_Title=" & Session("CertTitle") & "&vScore=" & vCertMark &  "&logo=" & svCustBanner

%>

