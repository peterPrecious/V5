<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vSponsorNo, vFirstName, vLastName, vEmail, vAccess, vAccessLink, vStatus, vMembMaxSponsor, vNoSponsors, vEligible

  vAccess     = "//" & svServer & "/" & Left(svCustId, 4) & fIf(svLang="FR", "-AP", "-SL")
  vAccessLink = "<a target='_blank' href='//" & svServer & "/" & Left(svCustId, 4) & fIf(svLang="FR", "-AP", "-SL") & "'>//" & svServer & "/" & Left(svCustId, 4) & fIf(svLang="FR", "-AP", "-SL") & "</a>"

  vSponsorNo = fDefault(Request("vSponsorNo"), svMembNo)

  sGetCust svCustId
  sGetMemb vSponsorNo

  vMembMaxSponsor    = vMemb_MaxSponsor '...store locally so this value doesn't get overwritten by the sponsors in list
  If vMembMaxSponsor = 0 Then vMembMaxSponsor = vCust_MaxSponsor
 
  If Request("vDelete").Count = 1 Then
    vMemb_No = Request("vDelete")
    sDeleteMemb
  End If

  If Request("vInactivate").Count = 1 Then
    sInactivateSponsor Request("vInactivate")
  End If

  If Request("vActivate").Count = 1 Then
    sActivateSponsor Request("vActivate")
  End If

  If Request("vExtend").Count = 1 Then
    sExtendSponsor Request("vExtend"), fFormatDate(DateAdd("d", 90, Request("vMembExpires")))
  End If


  If Request("vMaxSponsor").Count = 1 Then
    sMaxSponsor vSponsorNo, Request("vMaxSponsor")
  End If

  If Request("vForm").Count = 1 Then
    If vNoSponsors < vMembMaxSponsor Then
      sExtractSponsors
      sAddSponsors vFirstName, vLastName, vEmail, vSponsorNo
    End If
  End If

  sGetMemb vSponsorNo
  vMembMaxSponsor    = vMemb_MaxSponsor                             '...store locally so this value doesn't get overwritten by the sponsors in list
  If vMembMaxSponsor = 0 Then vMembMaxSponsor = vCust_MaxSponsor

  vNoSponsors = fNoSponsors (vSponsorNo)
  
  vEligible = fMax(vMembMaxSponsor - vNoSponsors, 0)
%>

<html>

<head>
  <title>Sponsors</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <% p0=vMembMaxSponsor %> <% p1=vAccessLink %>

  <div style="width:80%; margin:auto; text-align:left;">
    <h1><!--[[-->Setting Up and Managing Your Sponsored Learners<!--]]--></h1>
    <h3><!--[[-->Procedure for Setting Up Your Sponsored Learners:<!--]]--></h3>
    <ul>
      <li><!--[[-->Enter the learner name(s) and email addresses in the boxes below.<!--]]--></li>
      <li><!--[[-->Click the <b>Add</b> button each time you enter a name.<!--]]--></li>
      <li><!--[[-->You are allowed a maximum of <b>^0</b> learners at any time.<!--]]--></li>
      <li><!--[[-->Note that the system automatically assigns each learner a password.<!--]]--></li>
    </ul>
    <h3><!--[[-->Rules for Setting Up Your Sponsored Learners:<!--]]--></h3>
    <ul>
      <li><!--[[-->Each learner has 90 days access from time of registration.<!--]]--></li>
      <li><!--[[-->When a Sponsored Learner&#39;s access expires you can Add in another learner or extend that same learner&#39;s access for another 90 days.<!--]]--></li>
      <li><!--[[-->You can Inactivate Sponsored Learners (ie should they leave your organization) but once inactivated they cannot be reactivated and are counted as Sponsored Learners until their access expires.<!--]]--></li>
      <li><!--[[-->Please direct your Sponsored Learners to enter this service at <b>^1</b> where they can enter their assigned Password below.<!--]]--></li>
    </ul>
    <% If svMembLevel > 2 Then %>
    <h3><!--[[-->As a Facilitator, you have the ability to edit the Sponsored Learner list using the functions shown with yellow buttons:<!--]]--></h3>
    <ul>
      <li><!--[[-->You can increase or decrease the Maximum Sponsored Learners available to this particular Sponsor.<!--]]--></li>
      <li><!--[[-->You can <b>Delete</b> Sponsored Learners.<!--]]--></li>
      <li><!--[[-->You can <b>Reactivate</b> Sponsored Learners that were Inactivated by the Sponsor.<!--]]--></li>
    </ul>
    <% End If  %>

  </div>


  <table class="table">
    <tr>
      <td>
        <% If svMembLevel > 2 Then %>
        <p>&nbsp;</p>
        <% 
            Dim vSelected, vOption
            vOption = ""
            For i = 1 To 12
              vSelected = "" : If vMembMaxSponsor = i Then vSelected = "selected"
              vOption = vbCrLf & vOption & "<option " & vSelected & " value='" & i & "'>" & i & "</option>"
            Next
        %>
        <div style="text-align: center">
          <form method="POST" action="Sponsors.asp">
            <table class="table">
              <tr>
                <td style="text-align: center" class="c3">
                  <!--[[-->Maximum Sponsored Learners Allowed<!--]]-->
                  <select size="1" name="vMaxSponsor"><%=vOption%></select>
                  <input class="button" type="submit" value="<%="<!--{{-->Go<!--}}-->"%>" name="bGo" class="button00-adm">
                </td>
              </tr>
            </table>
            <input type="hidden" name="vSponsorNo" value="<%=vSponsorNo%>">
          </form>
        </div>
        <% End If %> 

      </td>
    </tr>
    <tr>
      <td style="text-align: center">
        <div style="text-align: center">
          <table class="table">
            <tr>
              <th colspan="9">
                <h1><!--[[-->Sponsored Learners List<!--]]--></h1>
                <h3><!--[[-->You can click on the Email field to send the Password to any Active Learner.<!--]]--></h3>
               </th>
            </tr>
            <tr>
              <th style="text-align: left; width:200px;"><!--[[-->Password<!--]]--></th>
              <th style="text-align: left; width:200px;"><!--[[-->First Name<!--]]--></th>
              <th style="text-align: left; width:200px;"><!--[[-->Last Name<!--]]--></th>
              <th style="text-align: left; width:200px;"><!--[[-->Email<!--]]--></th>
              <th style="text-align: center; width:200px;"><!--[[-->Added<!--]]--></th>
              <th style="text-align: center; width:200px;"><!--[[-->Expires<!--]]--></th>
              <th style="text-align: center; width:200px;"><!--[[--># Visits<!--]]--></th>
              <th style="text-align: center; width:200px;"><!--[[-->Status<!--]]--></th>
              <th style="text-align: center; width:200px;"><!--[[-->Action<!--]]--></th>
            </tr>
            <%
             sGetSponsors vSponsorNo
             Do While Not oRs.Eof
               sReadMemb
               vStatus = fIf(vMemb_Active, "Active", "Inactive")
               If vMemb_Active And vMemb_Expires <= Now Then vStatus = "Expired"

               If vMemb_Active And vMemb_Expires > Now Then 
                 vEmail = "<a href='mailto:" & vMemb_Email & "?subject=E-learning enrollment&body=Click " & vAccess & " and enter your password, which is: " & vMemb_Id  & ".  Enjoy.'>" & vMemb_Email & "</a>"
               Else
                 vEmail = vMemb_Email  
               End If
            %>
            <tr>
              <td style="white-space: nowrap"><%=vMemb_Id%></td>
              <td style="white-space: nowrap"><%=vMemb_FirstName%></td>
              <td style="white-space: nowrap"><%=vMemb_LastName%></td>
              <td style="white-space: nowrap"><%=vEmail%></td>
              <td style="white-space: nowrap; text-align: center"><%=fFormatDate(vMemb_FirstVisit)%></td>
              <td style="white-space: nowrap; text-align: center"><%=fFormatDate(vMemb_Expires)%></td>
              <td style="white-space: nowrap; text-align: center"><%=vMemb_NoVisits%></td>
              <td style="white-space: nowrap; text-align: center"><b><%=vStatus%></b></td>
              <td style="white-space: nowrap; text-align: center"><% If svMembLevel > 2 Then %>
                <input class="button" type="button" onclick="location.href = 'Sponsors.asp?vDelete=<%=vMemb_No%>&amp;vSponsorNo=<%=vSponsorNo%>'" value="<%="<!--{{-->Delete<!--}}-->"%>" name="bDelete" class="button085-adm">
                <% End If %> <% If vMemb_Active And vMemb_Expires > Now Then %>
                <input class="button" type="button" onclick="location.href = 'Sponsors.asp?vInactivate=<%=vMemb_No%>&amp;vSponsorNo=<%=vSponsorNo%>'" value="<%="<!--{{-->Inactivate<!--}}-->"%>" name="bInactivate" class="button085">
                <% ElseIf vMemb_Active And vMemb_Expires <= Now Then %>
                <input class="button" type="button" onclick="location.href = 'Sponsors.asp?vExtend=<%=vMemb_No%>&amp;vMembExpires=<%=vMemb_Expires%>&amp;vSponsorNo=<%=vSponsorNo%>'" value="<%="<!--{{-->Extend<!--}}-->"%>" name="bExtend" class="button085">
                <% ElseIf svMembLevel > 2 Then %>
                <input type="button" onclick="location.href = 'Sponsors.asp?vActivate=<%=vMemb_No%>&amp;vSponsorNo=<%=vSponsorNo%>'" value="<%="<!--{{-->Reactivate<!--}}-->"%>" name="bReactivate" class="button085-adm">
                <% End If %> 
              </td>
            </tr>
            <%
               oRs.MoveNext
             Loop
             sCloseDb
             Set oRs = Nothing
            %>
          </table>
          <p>&nbsp;</p>
        </div>
      </td>
    </tr>
    <% If vNoSponsors >= vMembMaxSponsor Then %>
    <tr>
      <td style="text-align: center">
        <h5><%p0=vNoSponsors%><!--[[-->You have already sponsored ^0 learners thus are not eligible to sponsor another learner at this time.<!--]]--></h5>
      </td>
    </tr>
    <%   If vNoSponsors = vMembMaxSponsor Then %>
    <tr>
      <td style="text-align: center">
        <h5><%p0=fFormatDate(fNextSponsorDate (vSponsorNo))%><!--[[-->On ^0 you can add another sponsored learner.<!--]]--></h5>
      </td>
    </tr>
    <%   End If %> 
    <% Else %>
    <tr>
      <td style="text-align: center">
        <div style="text-align: center">
          <form method="POST" action="Sponsors.asp">
            <table class="table">
              <tr>
                <td colspan="4" style="text-align: center">
                  <h1><!--[[-->Add a Sponsored Learner<!--]]--></h1>
                  <h3><%p0=vEligible%><!--[[-->You are eligible to add <b>^0</b> sponsored learner(s).&nbsp; Enter the fields below carefully as once you click <b>Add</b> they cannot be modified.<!--]]--></h3>
                </td>
              </tr>
              <tr>
                <th style="text-align: center; white-space: nowrap"><!--[[-->First Name<!--]]--></th>
                <th style="text-align: center; white-space: nowrap"><!--[[-->Last Name<!--]]--></th>
                <th style="text-align: center; white-space: nowrap"><!--[[-->Email Address<!--]]--></th>
                <th style="text-align: center; white-space: nowrap"><!--[[-->Action<!--]]--></th>
              </tr>
              <tr>
                <td style="text-align: center"><input type="text" size="20" name="vFirstName"></td>
                <td style="text-align: center"><input type="text" size="20" name="vLastName"></td>
                <td style="text-align: center"><input type="text" size="36" name="vEmail"></td>
                <td style="text-align: center"><input type="submit" value="<%="<!--{{-->Add<!--}}-->"%>" name="bAdd" class="button" tabindex="4"></td>
              </tr>
            </table>
            <input type="hidden" name="vForm" value="y">
            <input type="hidden" name="vSponsorNo" value="<%=vSponsorNo%>">
          </form>
        </div>
      </td>
    </tr>
    <% End If %>
    <tr>
      <td style="text-align: center">&nbsp;
        <p>
          <input type="button" onclick="location.href = 'Info.asp'" value="<%="<!--{{-->Info Page<!--}}-->"%>" name="bInfo" class="button">
          <% If svMembLevel > 2 Then %>
          <input type="button" onclick="location.href = 'User.asp?vMembNo=<%=vSponsorNo%>'" value="<%="<!--{{-->Sponsor Profile<!--}}-->"%>" name="bSponsor" class="button">
          <input type="button" onclick="location.href = 'Users_o.asp?vLastValue=<%=svMembLastName & svMembFirstName%>'" value="<%="<!--{{-->Learner List<!--}}-->"%>" name="bLearnerList" class="button">
          <% End If %>
        </p>
      </td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
