<!--#include virtual = "V5/Inc/Setup.asp"-->
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
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
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
        <h3><b><!--[[-->Status<!--]]--> : <%=fModStatus(svMembNo, vMods_Id)%></b></h3>


        <% 
          '...display this note if there are assessment activities
          If vCust_ResetStatus > 0 And vNoAttempts > 0 Then 
            p1 = vCust_ResetStatus
        %>
        <p><!--[[-->Reflects assessment activities during the past ^1 days.<!--]]--></p>
        <% 
          End If 
        %>

        <table border="0" cellspacing="0" cellpadding="2" id="table5">

          <% If vNoAttempts > 0 Then%>
          <tr>
            <td><b><!--[[-->Best assessment score<!--]]--> : </b></td>
            <td class="d2">
            <%=vBestScore%>

            <% If vBestScore >= vPassScore And vMods_VuCert Then %>
              <a <%=fstatx%> class="d2" href="javascript:fullScreen('<%=fCertificateUrl("", "", vBestScore, vLastScore, vMods_Id, vMods_Title, "", "", "", vProg_Id, "", "", "")%>')"><!--[[-->Certificate<!--]]--></a>
            <% End If %>

            </td>
          </tr>
          <tr>
            <td><b><!--[[-->Number of attempts<!--]]--> : </b></td>
            <td><%=vNoAttempts%></td>
          </tr>
          <tr>
            <td><b><!--[[-->Last assessment attempt<!--]]--> : </b></td>
            <td><%=fFormatDate(vLastScore)%></td>
          </tr>
          <% End If %>


          <%
            vTimeSpent = fTimeSpent(svMembNo, vMods_Id)
            If vTimeSpent > 0 Then
          %>
          <tr>
            <td><b><!--[[-->Last time module was reviewed<!--]]--> : </b></td>
            <td><%=fLastSpent(svMembNo, vMods_Id)%></td>
          </tr>
          <tr>
            <td><b><!--[[-->Total time spent in module<!--]]--> : </b></td>
            <td><%=fTimeSpent(svMembNo, vMods_Id)%>&nbsp;<!--[[-->minutes<!--]]--></td>
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
            <td><p class="c3"><!--[[-->Status Legend<!--]]--></p></td>
          </tr>
          <tr>
            <td><!--[[--><b>Not Started</b>: Module has not been accessed.<!--]]--></td>
          </tr>
          <tr>
            <td><!--[[--><b>Reviewed</b>: Module has been previously accessed and may be in progress or completed.<!--]]--></td>
          </tr>
          <tr>
            <td><!--[[--><b>Completed</b>: Note: Only available on some modules.  Some modules have a special assessment that, when passed, will automatically change the status to &quot;Completed&quot;.<!--]]-->
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