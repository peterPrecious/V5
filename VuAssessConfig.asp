<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Code/ModuleStatusRoutines.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<%
  Dim vProgId,vModId, vResponse, vNoAttempts

  If Not svSecure Then 

    vResponse = "invalid"

  Else

    vProgId = Ucase(Request("vProgId"))
    vModId  = Ucase(Request("vModId"))
  
    sGetProg (vProgId)  '...can be invalid
    sGetMods (vModId)   '...can be invalid
    sGetCust (svCustId) '...will always be valid

    vResponse = "memb_level=" & svMembLevel

    vResponse = vResponse & "&language=" & Right(vModId, 2)

    vResponse = vResponse & "&best_score=" & fBestScore (svMembNo, vModId)

    If Len(vProg_AssessmentCert) > 0 Then
      vResponse = vResponse & "&custom_cert=" & vProg_AssessmentCert        
    ElseIf Len(vCust_AssessmentCert) > 0 Then
      vResponse = vResponse & "&custom_cert=" & vCust_AssessmentCert        
    End If

    If vProg_AssessmentAttempts > 0 Then
      vResponse = vResponse & "&max_attempts=" & vProg_AssessmentAttempts
    ElseIf vCust_AssessmentAttempts > 0 Then
      vResponse = vResponse & "&max_attempts=" & vCust_AssessmentAttempts
    Else      
      vResponse = vResponse & "&max_attempts=" & 3
    End If

    '...if empty assume default (80%), if .01 allow 0
    If vProg_AssessmentScore = 0.01 Then
      vResponse = vResponse & "&pass_grade=0"
    ElseIf vProg_AssessmentScore > 0 Then
      vResponse = vResponse & "&pass_grade=" & vProg_AssessmentScore * 100
    ElseIf vCust_AssessmentScore = 0.01 Then
      vResponse = vResponse & "&pass_grade=0"
    ElseIf vCust_AssessmentScore > 0 Then
      vResponse = vResponse & "&pass_grade=" & vCust_AssessmentScore * 100      
    Else
      vResponse = vResponse & "&pass_grade=" & 80      
    End If

    vNoAttempts = fNoAttempts(svMembNo, vModId)
    vResponse = vResponse & "&num_attempts=" & vNoAttempts

    If Len(svCustBanner) > 0 Then vResponse = vResponse & "&logo=" & svCustBanner

  End If
     
  Response.Write vResponse

%>








