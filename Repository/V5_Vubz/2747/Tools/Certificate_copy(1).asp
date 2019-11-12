<!--#include virtual = "V5\Inc\Setup.asp"-->
<% vClose = "Y" %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5\Inc\Db_Cust.asp"-->
<!--#include virtual = "V5\Inc\Db_Prog.asp"-->
<!--#include virtual = "V5\Inc\Db_Phra.asp"-->
<!--#include virtual = "V5\Inc\Db_Memb.asp"-->
<!--#include virtual = "V5\Inc\Db_Logs.asp"-->
<!--#include virtual = "V5\Inc\Certificate.asp"-->
<!--#include virtual = "V5\Code\ModuleStatusRoutines.asp"-->

<%
  Dim vOk, vErrMsg, vMod, v1, v2, v3, v4, v5, aMemo, vInst, vCourse, vLastScore, vScores, vUrl

  vOk = True
  vLastScore = "Jan 1, 2000"

  '...check quiz was taken
  If vOk Then
    vMod = "9427" & svLang
    v1 = fFirstScore (svMembNo, vMod)
    If v1 > 0 Then
      vLastScore = fMax(cDate(fLastScore (svMembNo, vMod)), cDate(vLastScore))
    Else
      vOk = False 
    End If
  End If  

  '...check quiz was taken
  If vOk Then
    vMod = "9495" & svLang
    v2 = fFirstScore (svMembNo, vMod)
    If v2 > 0 Then
      vLastScore = fMax(cDate(fLastScore (svMembNo, vMod)), cDate(vLastScore))
    Else
      vOk = False 
    End If
  End If  

  '...check quiz was taken
  If vOk Then
    vMod = "9497" & svLang
    v3 = fFirstScore (svMembNo, vMod)
    If v3 > 0 Then
      vLastScore = fMax(cDate(fLastScore (svMembNo, vMod)), cDate(vLastScore))
    Else
      vOk = False 
    End If
  End If  


  '...check quiz was taken
  If vOk Then
    vMod = "9498" & svLang
    v4 = fFirstScore (svMembNo, vMod)
    If v4 > 0 Then
      vLastScore = fMax(cDate(fLastScore (svMembNo, vMod)), cDate(vLastScore))
    Else
      vOk = False 
    End If
  End If  


  '...check if survey has been taken
'  If vOk Then
'    vMod = "9550" & svLang
'    vLastScore = fSurveyCompleted (svMembNo, vMod)
'    If Not IsDate(vLastScore) Then
'      vOk = False 
'    End If
'  End If  



  '...set to true for debugging/cert else comment out
  vOk = True
 
  If Not vOk Then
    If svLang = "FR" Then
	    vErrMsg = "Afin d'obtenir un certificat d'achèvement, vous devez remplir l'ensemble des quatre tests, plus l'enquête."
    Else
      vErrMsg = "In order to be granted a Certificate of Completion you must complete all four tests plus the survey."
    End If  
    
    Response.Redirect "/V5/Code/Error.asp?vErr=" & Server.UrlEncode(vErrMsg)

  Else
  
    '...provide 5 scores
'   vScores = Right("000" & cInt(v1/4), 3) & Right("000" & cInt(v2/4), 3) & Right("000" & cInt(v3/4), 3) & Right("000" & cInt(v4/4), 3) & Right("000" & v5, 3)
    vScores = cInt(v1/4) & "|" & cInt(v2/4) & "|" & cInt(v3/4) & "|" & cInt(v4/4)
    v5 = cInt(v1/4) + cInt(v2/4) + cInt(v3/4) + cInt(v4/4)
    
    '...pass the institution and course info to cert
    sGetMemb (svMembNo)

    vMemb_Memo					        = oRs("Memb_Memo")
    aMemo                       = Split(vMemb_Memo, "|")
    If Ubound(aMemo) < 5 Then 
      vMemb_Memo                = vMemb_Memo & "||||"
      aMemo                     = Split(vMemb_Memo, "|")
    End If


'    If Len(vMemb_Memo) = 0 Then 
'      vInst   = "University of Western Ontario"
'      vCourse = "Business 101"
'    Else
'      aMemo = Split(vMemb_Memo, "|")
      vInst   = aMemo(1)
      vCourse = aMemo(4)
'    End If












'   vUrl = "/V5/Assessments/CustomCerts/MEVT_" & svLang & "/Default.htm?vFirstName=" & svMembFirstName & "&vLastName=" & svMembLastName & "&vLastScore=" & fFormatDate(vLastScore) & "&vMods_Id=0000EN" & "&vScores=" & vScores & "&vInst=" & vInst & "&vCourse=" & vCourse
    vUrl = fCertificateUrl("", "", v5, fFormatDate(vLastScore), "", "", "", "", "", "", "", vScores & "|" & vInst & "|" & vCourse)

    Response.Redirect vUrl

  End If
%>