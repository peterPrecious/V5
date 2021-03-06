﻿<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Certificate.asp"-->
<!--#include file = "ModuleStatusRoutines.asp"-->

<%  
  Dim vNoAttempts, vCompleted, vBestScore, vLastScore, vCertUrl, vPassScore, vCustomCert, vTimeSpent

  vMods_Id    = Request("vModId")
  vProg_Id    = Request("vProgId")

  sGetMods (vMods_Id)
  sGetCust (svCustId)
  sGetProg (vProg_Id)

  '..what is the passing grade
  If vProg_AssessmentScore = .01 Then
    vPassScore = .01
  ElseIf vProg_AssessmentScore > .01 Then
    vPassScore = vProg_AssessmentScore    
  ElseIf vCust_AssessmentScore = .01 Then
    vPassScore = 0
  ElseIf vCust_AssessmentScore > .01 Then
    vPassScore = vCust_AssessmentScore    
  Else
    vPassScore = .8
  End If
  vPassScore = vPassScore * 100

  '...use a custom cert?
  If Len(vProg_AssessmentCert) > 0 Then
    vCustomCert = vProg_AssessmentCert
  ElseIf Len(vCust_AssessmentCert) > 0 Then
    vCustomCert = vCust_AssessmentCert
  Else
    vCustomCert = ""
  End If
  
  vNoAttempts = fNoAttempts(svMembNo, vMods_Id)
  vBestScore  = fBestScore(svMembNo, vMods_Id)
  vLastScore  = fLastScore(svMembNo, vMods_Id)
  vCompleted  = fIsCompleteMod(svMembNo, vMods_Id)
%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script src="/V5/Inc/Launch.js"></script>
  <title>StatusModules</title>
</head>

<body>

  <% Server.Execute vShellHi %>
  <form>
    <table border="0" id="table4" style="border-collapse: collapse" bordercolor="#DDEEF9" width="100%" cellpadding="2">
      <tr>
        <td align="center">
        <h1><%=vMods_Title%></h1>
        </td>
      </tr>
      <tr>
        <td valign="top" align="center">
        <h3><b><!--webbot bot='PurpleText' PREVIEW='Status'--><%=fPhra(000244)%> : <%=fModStatus(svMembNo, vMods_Id)%></b></h3>


        <% 
          '...display this note if there are assessment activities
          If vCust_ResetStatus > 0 And vNoAttempts > 0 Then 
            p1 = vCust_ResetStatus
        %>
        <p><!--webbot bot='PurpleText' PREVIEW='Reflects assessment activities during the past ^1 days.'--><%=fPhra(000511)%></p>
        <% 
          End If 
        %>

        <table border="0" cellspacing="0" cellpadding="2" id="table5">

          <% If vNoAttempts > 0 Then%>
          <tr>
            <td><b><!--webbot bot='PurpleText' PREVIEW='Best assessment score'--><%=fPhra(000079)%> : </b></td>
            <td class="d2">
            <%=vBestScore%>

            <% If vBestScore >= vPassScore And vMods_VuCert Then %>
              <a <%=fstatx%> class="d2" href="javascript:fullScreen('<%=fCertificateUrl("", "", vBestScore, vLastScore, vMods_Id, vMods_Title, "", "", "", vProg_Id, "", "", "")%>')"><!--webbot bot='PurpleText' PREVIEW='Certificate'--><%=fPhra(000089)%></a>
            <% End If %>

            </td>
          </tr>
          <tr>
            <td><b><!--webbot bot='PurpleText' PREVIEW='Number of attempts'--><%=fPhra(000197)%> : </b></td>
            <td><%=vNoAttempts%></td>
          </tr>
          <tr>
            <td><b><!--webbot bot='PurpleText' PREVIEW='Last assessment attempt'--><%=fPhra(000162)%> : </b></td>
            <td><%=fFormatDate(vLastScore)%></td>
          </tr>
          <% End If %>


          <%
            vTimeSpent = fTimeSpent(svMembNo, vMods_Id)
            If vTimeSpent > 0 Then
          %>
          <tr>
            <td><b><!--webbot bot='PurpleText' PREVIEW='Last time module was reviewed'--><%=fPhra(000590)%> : </b></td>
            <td><%=fLastSpent(svMembNo, vMods_Id)%></td>
          </tr>
          <tr>
            <td><b><!--webbot bot='PurpleText' PREVIEW='Total time spent in module'--><%=fPhra(000028)%> : </b></td>
            <td><%=fTimeSpent(svMembNo, vMods_Id)%>&nbsp;<!--webbot bot='PurpleText' PREVIEW='minutes'--><%=fPhra(000174)%></td>
          </tr>
          <%
            End If 
          %>

        </table>
        </td>
      </tr>
      <tr>
        <td>
          <br>&nbsp;
          <table class="table">
          <tr>
            <td><p class="c3"><!--webbot bot='PurpleText' PREVIEW='Status Legend'--><%=fPhra(000248)%></p></td>
          </tr>
          <tr>
            <td><!--webbot bot='PurpleText' PREVIEW='<b>Not Started</b>: Module has not been accessed.'--><%=fPhra(000057)%></td>
          </tr>
          <tr>
            <td><!--webbot bot='PurpleText' PREVIEW='<b>Reviewed</b>: Module has been previously accessed and may be in progress or completed.'--><%=fPhra(000058)%></td>
          </tr>
          <tr>
            <td><!--webbot bot='PurpleText' PREVIEW='<b>Completed</b>: Note: Only available on some modules.  Some modules have a special assessment that, when passed, will automatically change the status to &quot;Completed&quot;.'--><%=fPhra(000056)%>
            </td>
          </tr>
        </table>
        </td>
      </tr>
    </table>
  </form>

  <h1><input onclick="javascript: window.close()" type="button" value="<%=fIf(svLang="FR", "Fermer", "Close")%>" name="bClose" class="button"></h1>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

