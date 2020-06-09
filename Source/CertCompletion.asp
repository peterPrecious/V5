<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vClose = "Y" %>
<!--#include virtual = "V5/Inc/ProgramStatusRoutines.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<%
  Dim vProgId, vProgCertType, vShowCert, vProgModTimeConstraint, vProgModScoreConstraint, vProgModAttemptConstraint, aProgMods, vErrMess, oRsCheck
 
  vProgId = fDefault(Request("vProgId"), Session("Ecom_Prog"))

  Session("CertType")     = "Completion"
  Session("CertId")       = vProgId 
  Session("CertMark")     = "-1"

  If Len(Request("vCertTitle")) > 0 Then
    Session("CertTitle")  = Request("vCertTitle")
  Else
    Session("CertTitle")  = fProgTitle(vProgId)
  End If

  Session("CertSample")   = ""   '...ensure certificate is NOT a sample
  Session("CertDate")     = Now()

  '...need to check Program Table to determine requirements (Prog_Cert)
  '   0 = Do NOT offer a certificate at the end of this program
  '   1 = Offer a certificate at the end of this program without constraints
  '   2 = Offer if at least X (Prog_CertTimeSpent) minutes spent in each module
  '   3 = Offer if all Tests/Sas attained at least X% (Prog_CertTestScore) within Y (Prog_CertTestAttempts) attempts 

  sOpenDbBase
  vSql = "SELECT * FROM Prog WHERE Prog_Id='" & vProgId & "'"
' sDebug
  Set oRsCheck = oDbBase.Execute(vSql)

  '...grab list of Modules in the current Prog
  aProgMods = Split(oRsCheck("Prog_Mods")," ")

  '...grab Certificate condition for current Prog
  vProgCertType = oRsCheck("Prog_Cert")
  If Len(vProgCertType) = 0 Then vProgCertType = 0

  vShowCert = False

  Select Case vProgCertType
    '...never
    Case 0
    '...always
    Case 1
      vShowCert = True
    '...only if total min met
    Case 2
      vProgModTimeConstraint = oRsCheck("Prog_CertTimeSpent")
      '...if req is 0, then no time constraint...but must have accessed EACH module once
      If vProgModTimeConstraint = 0 Then vProgModTimeConstraint = 1
      If fProgModsTimeSpentMet(svMembNo,vModId,aProgMods,vProgModTimeConstraint) Then 
        vShowCert = True
      Else
        If oRsCheck("Prog_CertTimeSpent") = 0 Then
          vErrMess = "In order to be granted a Certificate of Completion, you need to review all modules in this program."
        Else
          vErrMess = "In order to be granted a Certificate of Completion, you need to properly review all modules in this program."
        End If
      End If
    '...if all Mod's score within set attemps met
    Case 3
      vProgModScoreConstraint = oRsCheck("Prog_CertTestScore")
      vProgModAttemptConstraint = oRsCheck("Prog_CertTestAttempts")
      '...if both values 0, then no score constraint...but must attempted test once
      If vProgModScoreConstraint = 0 and vProgModAttemptConstraint = 0 Then 
        vProgModAttemptConstraint = 1
      End If
      If fProgModsScoreAttemptMet(svMembNo,aProgMods,vProgModScoreConstraint,vProgModAttemptConstraint) Then 
        vShowCert = True
      Else
        If oRsCheck("Prog_CertTestScore") = 0 And oRsCheck("Prog_CertTestAttempts") = 0 Then
          vErrMess = "In order to be granted a Certificate of Completion, you need to attempt each self assessment that is offered at the end of each module."
        Else
          vErrMess = "In order to be granted a Certificate of Completion, you need to attain at least " & vProgModScoreConstraint & " percent in each assessment"
          If vProgModAttemptConstraint > 0 Then 
            vErrMess = vErrMess & " within " & vProgModAttemptConstraint & " attempts"
          End If
          vErrMess = vErrMess & "."
        End If
      End If
  End Select

  oRsCheck.Close
  Set oRsCheck = Nothing

  sCloseDBBase

  If vShowCert = True Then
    Response.Redirect "Certificate.asp"
  Else
    Response.Redirect "Error.asp?vClose=Y&vErr=" & Server.HtmlEncode(vErrMess) & "&vReturn=javascript:close()"
  End If








  '...find total Time Spent in each item in Module List by Program (by a given user)
  Function fProgModsTimeSpentMet(vMembNo, vProgId, aProgMods, vTimeConstraint)
    '...access each Module in provided list to ensure that X amount of time spent
    Dim vSql, oRs2, vMod

    fProgModsTimeSpentMet = True

    sOpenDb2
    For Each vMod in aProgMods
      vSql = "SELECT Logs_Item FROM Logs WITH (nolock) WHERE (Logs_MembNo = " & vMembNo & ") AND (Logs_Type = 'P') AND (LEFT(Logs_Item, 14) = '" & vProgId & "|" & vMod & "')"
      Set oRs2 = oDb2.Execute(vSql)
      If oRs2.Eof Then 
        fProgModsTimeSpentMet = False
        Exit For
      ElseIf Cint(Right(oRs2("Logs_Item"), 6)) < vTimeConstraint Then 
        fProgModsTimeSpentMet = False
        Exit For
      End If
    Next

    sCloseDb2
    Set oRs2 = Nothing
  End Function

  '...find highest Score within set Attempts in each item in Module List by Program (by a given user)
  Function fProgModsScoreAttemptMet(vMembNo, aProgMods, vScoreConstraint, vAttemptConstraint)

    '...access each Module in provided list to ensure that X% acheived in Y attempts
    Dim vSql, oRs2, vMod, vScoreBest, vScoreCount
    fProgModsScoreAttemptMet = True
    sOpenDb2
    For Each vMod in aProgMods
      '...ensure number of attempts do not excede limit 
      If vAttemptConstraint > 0 Then  
        vSql = "SELECT COUNT(Logs_Item) AS ScoreCount FROM Logs WITH (nolock) WHERE (Logs_Type = 'T') AND (Logs_MembNo = " & vMembNo & ") AND (LEFT(Logs_Item, 6) = '" & vMod & "') AND (LEN(Logs_Item) = 10)"
        Set oRs2 = oDb2.Execute(vSql)
        vScoreCount = fDefault(oRs2("ScoreCount"), 0)
        If vScoreCount > vAttemptConstraint Then 
          fProgModsScoreAttemptMet = False
          Exit For
        End If
      End If
      '...get high score
      vSql = "SELECT MAX(RIGHT(Logs_Item, 3)) AS HighScore FROM Logs WITH (nolock) WHERE (Logs_Type = 'T') AND (Logs_MembNo = " & vMembNo & ") AND (LEFT(Logs_Item, 6) = '" & vMod & "') AND (LEN(Logs_Item) = 10)"
      Set oRs2 = oDb2.Execute(vSql)
      If oRs2.Eof Then 
        fProgModsScoreAttemptMet = False
        Exit For
      Else
        vScoreBest = fDefault(oRs2("HighScore"), 0)
        If vScoreBest < vScoreConstraint Then 
          fProgModsScoreAttemptMet = False
          Exit For
        End If
      End If

    Next

    sCloseDb2
    Set oRs2 = Nothing
  End Function












%>