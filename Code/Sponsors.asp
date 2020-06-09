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
    <h1><!--webbot bot='PurpleText' PREVIEW='Setting Up and Managing Your Sponsored Learners'--><%=fPhra(000526)%></h1>
    <h3><!--webbot bot='PurpleText' PREVIEW='Procedure for Setting Up Your Sponsored Learners:'--><%=fPhra(000527)%></h3>
    <ul>
      <li><!--webbot bot='PurpleText' PREVIEW='Enter the learner name(s) and email addresses in the boxes below.'--><%=fPhra(000528)%></li>
      <li><!--webbot bot='PurpleText' PREVIEW='Click the <b>Add</b> button each time you enter a name.'--><%=fPhra(000529)%></li>
      <li><!--webbot bot='PurpleText' PREVIEW='You are allowed a maximum of <b>^0</b> learners at any time.'--><%=fPhra(000530)%></li>
      <li><!--webbot bot='PurpleText' PREVIEW='Note that the system automatically assigns each learner a password.'--><%=fPhra(000531)%></li>
    </ul>
    <h3><!--webbot bot='PurpleText' PREVIEW='Rules for Setting Up Your Sponsored Learners:'--><%=fPhra(000532)%></h3>
    <ul>
      <li><!--webbot bot='PurpleText' PREVIEW='Each learner has 90 days access from time of registration.'--><%=fPhra(000533)%></li>
      <li><!--webbot bot='PurpleText' PREVIEW='When a Sponsored Learner&#39;s access expires you can Add in another learner or extend that same learner&#39;s access for another 90 days.'--><%=fPhra(000534)%></li>
      <li><!--webbot bot='PurpleText' PREVIEW='You can Inactivate Sponsored Learners (ie should they leave your organization) but once inactivated they cannot be reactivated and are counted as Sponsored Learners until their access expires.'--><%=fPhra(000535)%></li>
      <li><!--webbot bot='PurpleText' PREVIEW='Please direct your Sponsored Learners to enter this service at <b>^1</b> where they can enter their assigned Password below.'--><%=fPhra(000536)%></li>
    </ul>
    <% If svMembLevel > 2 Then %>
    <h3><!--webbot bot='PurpleText' PREVIEW='As a Facilitator, you have the ability to edit the Sponsored Learner list using the functions shown with yellow buttons:'--><%=fPhra(000537)%></h3>
    <ul>
      <li><!--webbot bot='PurpleText' PREVIEW='You can increase or decrease the Maximum Sponsored Learners available to this particular Sponsor.'--><%=fPhra(000538)%></li>
      <li><!--webbot bot='PurpleText' PREVIEW='You can <b>Delete</b> Sponsored Learners.'--><%=fPhra(000539)%></li>
      <li><!--webbot bot='PurpleText' PREVIEW='You can <b>Reactivate</b> Sponsored Learners that were Inactivated by the Sponsor.'--><%=fPhra(000540)%></li>
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
                  <!--webbot bot='PurpleText' PREVIEW='Maximum Sponsored Learners Allowed'--><%=fPhra(000498)%>
                  <select size="1" name="vMaxSponsor"><%=vOption%></select>
                  <input class="button" type="submit" value="<%=fPhraH(000432)%>" name="bGo" class="button00-adm">
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
                <h1><!--webbot bot='PurpleText' PREVIEW='Sponsored Learners List'--><%=fPhra(000509)%></h1>
                <h3><!--webbot bot='PurpleText' PREVIEW='You can click on the Email field to send the Password to any Active Learner.'--><%=fPhra(000499)%></h3>
               </th>
            </tr>
            <tr>
              <th style="text-align: left; width:200px;"><!--webbot bot='PurpleText' PREVIEW='Password'--><%=fPhra(000211)%></th>
              <th style="text-align: left; width:200px;"><!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%></th>
              <th style="text-align: left; width:200px;"><!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%></th>
              <th style="text-align: left; width:200px;"><!--webbot bot='PurpleText' PREVIEW='Email'--><%=fPhra(000342)%></th>
              <th style="text-align: center; width:200px;"><!--webbot bot='PurpleText' PREVIEW='Added'--><%=fPhra(000500)%></th>
              <th style="text-align: center; width:200px;"><!--webbot bot='PurpleText' PREVIEW='Expires'--><%=fPhra(000137)%></th>
              <th style="text-align: center; width:200px;"><!--webbot bot='PurpleText' PREVIEW='# Visits'--><%=fPhra(000501)%></th>
              <th style="text-align: center; width:200px;"><!--webbot bot='PurpleText' PREVIEW='Status'--><%=fPhra(000244)%></th>
              <th style="text-align: center; width:200px;"><!--webbot bot='PurpleText' PREVIEW='Action'--><%=fPhra(000502)%></th>
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
                <input class="button" type="button" onclick="location.href = 'Sponsors.asp?vDelete=<%=vMemb_No%>&amp;vSponsorNo=<%=vSponsorNo%>'" value="<%=fPhraH(000494)%>" name="bDelete" class="button085-adm">
                <% End If %> <% If vMemb_Active And vMemb_Expires > Now Then %>
                <input class="button" type="button" onclick="location.href = 'Sponsors.asp?vInactivate=<%=vMemb_No%>&amp;vSponsorNo=<%=vSponsorNo%>'" value="<%=fPhraH(000495)%>" name="bInactivate" class="button085">
                <% ElseIf vMemb_Active And vMemb_Expires <= Now Then %>
                <input class="button" type="button" onclick="location.href = 'Sponsors.asp?vExtend=<%=vMemb_No%>&amp;vMembExpires=<%=vMemb_Expires%>&amp;vSponsorNo=<%=vSponsorNo%>'" value="<%=fPhraH(000496)%>" name="bExtend" class="button085">
                <% ElseIf svMembLevel > 2 Then %>
                <input type="button" onclick="location.href = 'Sponsors.asp?vActivate=<%=vMemb_No%>&amp;vSponsorNo=<%=vSponsorNo%>'" value="<%=fPhraH(000497)%>" name="bReactivate" class="button085-adm">
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
        <h5><%p0=vNoSponsors%><!--webbot bot='PurpleText' PREVIEW='You have already sponsored ^0 learners thus are not eligible to sponsor another learner at this time.'--><%=fPhra(000503)%></h5>
      </td>
    </tr>
    <%   If vNoSponsors = vMembMaxSponsor Then %>
    <tr>
      <td style="text-align: center">
        <h5><%p0=fFormatDate(fNextSponsorDate (vSponsorNo))%><!--webbot bot='PurpleText' PREVIEW='On ^0 you can add another sponsored learner.'--><%=fPhra(000504)%></h5>
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
                  <h1><!--webbot bot='PurpleText' PREVIEW='Add a Sponsored Learner'--><%=fPhra(000505)%></h1>
                  <h3><%p0=vEligible%><!--webbot bot='PurpleText' PREVIEW='You are eligible to add <b>^0</b> sponsored learner(s).&nbsp; Enter the fields below carefully as once you click <b>Add</b> they cannot be modified.'--><%=fPhra(000549)%></h3>
                </td>
              </tr>
              <tr>
                <th style="text-align: center; white-space: nowrap"><!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%></th>
                <th style="text-align: center; white-space: nowrap"><!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%></th>
                <th style="text-align: center; white-space: nowrap"><!--webbot bot='PurpleText' PREVIEW='Email Address'--><%=fPhra(000126)%></th>
                <th style="text-align: center; white-space: nowrap"><!--webbot bot='PurpleText' PREVIEW='Action'--><%=fPhra(000502)%></th>
              </tr>
              <tr>
                <td style="text-align: center"><input type="text" size="20" name="vFirstName"></td>
                <td style="text-align: center"><input type="text" size="20" name="vLastName"></td>
                <td style="text-align: center"><input type="text" size="36" name="vEmail"></td>
                <td style="text-align: center"><input type="submit" value="<%=fPhraH(000506)%>" name="bAdd" class="button" tabindex="4"></td>
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
          <input type="button" onclick="location.href = 'Info.asp'" value="<%=fPhraH(000160)%>" name="bInfo" class="button">
          <% If svMembLevel > 2 Then %>
          <input type="button" onclick="location.href = 'User.asp?vMembNo=<%=vSponsorNo%>'" value="<%=fPhraH(000507)%>" name="bSponsor" class="button">
          <input type="button" onclick="location.href = 'Users_o.asp?vLastValue=<%=svMembLastName & svMembFirstName%>'" value="<%=fPhraH(000508)%>" name="bLearnerList" class="button">
          <% End If %>
        </p>
      </td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


