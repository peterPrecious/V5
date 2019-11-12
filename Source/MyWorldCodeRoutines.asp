<%

  Function fVuAssessLink (vModId, vLinkTitle, vCertTitle)

    Dim i, j, k, vAssessmentScore, vBestScore, vFolder


  if vModId = "40374EN" Then Stop
 ' stop

    '...get VuAssess Title
    fVuAssessLink = vLinkTitle
    If Len(vLinkTitle) = 0  Then fVuAssessLink = vCertTitle
    If Len(fVuAssessLink) = 0 Then fVuAssessLink = Server.HtmlEncode("<!--{{-->Examination<!--}}-->")
    vAlt = fVuAssessLink

    '...find passing score requirement and if met then generate a cert
    vAssessmentScore = vCust_AssessmentScore     
    If vAssessmentScore = 0 Then vAssessmentScore = vProg_AssessmentScore     
    If vAssessmentScore = 0 Then vAssessmentScore = .80

    vBestScore = fBestScore (svMembNo, vModId)/100 
    If vBestScore >= vAssessmentScore Then 

      '...assuming they want a cert
      If vMods_VuCert Then
        If Len(vProg_AssessmentCert) > 0 Then
         vFolder = vProg_AssessmentCert
        ElseIf Len(vCust_AssessmentCert) > 0 Then
         vFolder = vCust_AssessmentCert
        Else
         vFolder = svLang
        End If
        i = fLastPassed(vModId, vProg_AssessmentScore)
  			vUrl   = "javascript:fullScreen('" & fCertificateUrl("", "", vBestScore * 100 , i, vModId, vCertTitle, "", "", "", vProg_Id, "", "", "") & "')"
   			fVuAssessLink = "<span class='c2'><a " & fStatX & " href=""" & vUrl & """>" & vAlt & "</a></span>"

      '...else just show the title without a link
      Else
        vUrl = ""
  			fVuAssessLink = "<span class='c2'>" & fVuAssessLink & "</span>"     
      End If

    '...if they haven't passed then allow another launch until they excede their max attempts
    Else

      If vProg_AssessmentAttempts > 0 Then
        vAttempts = vProg_AssessmentAttempts
      ElseIf vCust_AssessmentAttempts > 0 Then
        vAttempts = vCust_AssessmentAttempts
      Else
        vAttempts = 3
      End If

      If vAttempts = 99 Or fNoAttempts(svMembNo, vModId) < vAttempts Then 
        vUrlTitle1 = Server.HtmlEncode("<!--{{-->Click here to launch the Assessment<!--}}-->")

        If Lcase(vMods_Type) = "fx" Then
'         vUrl = "javascript:" & vMods_Script & "('" & vProg_Id & "|" & vModId & "|" & vProg_Test & "|" & vProg_Bookmark & "|" & vProg_CompletedButton & "')"
'         vUrl = "/V5/LaunchObjects.asp?vModId=" & vMods_Id & "&vNext=" & svPage 
          vUrl = "/V5/LaunchObjects.asp?vModId=" & vProg_Id & "|" & vMods_Id & "&vNext=MyWorld.asp" 
        Else
          vUrl = "javascript:" & vMods_Script & "('" & vProg_Id & "|" & vModId & "|" & vProg_Test & "|" & vProg_Bookmark & "|" & vProg_CompletedButton & "')"
        End If

        fVuAssessLink = "<span class='c2'><a " & fStatX & " href=""" & vUrl & """ title='" & vUrlTitle1 & "'>" & fVuAssessLink & "</a></span>"

      Else 
        vUrl = "javascript:fAlert()"
        vUrlTitle1 = Server.HtmlEncode("<!--{{-->You have no more attempts available for this assessment.!--}}-->")
        fVuAssessLink = "<span class='c2'><a " & fStatX & " href=""" & vUrl & """ title='" & vUrlTitle1 & "'>" & fVuAssessLink & "</a></span>"
      End If 


    End If

    '...add in the status
    fVuAssessLink = fVuAssessLink & "&nbsp;<span class='green'>[" & fAssessmentStatus (svMembNo, vModId) & "]</span>"

  End Function








%>