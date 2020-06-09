<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<% 'stop  %>

<%
  Dim vMessage, vPrograms, vNext, vMemb_ProgramsLength, vAlert, bAlertOk, vAlertMsg, vMembId, bNoEditProfile
  
  vNext    = fDefault(Request("vNext"), "Users_O.asp?vSort=id&vStart=" & vMemb_Id)
  vMessage = Request.QueryString("vMessage")

  '...used in translation engine to Id Type
  p0 = fIf(svCustPwd, fPhraH(000411), fPhraH(000211))

  '...turn this on if the parent site offers email alerts to their G2 sites
  bAlertOk = fCustParentG2alertOk (svCustAcctId) 

  vAlert = Request("vAlert")
  If vAlert <> "" Then
    sUpdateCustG2alert svCustAcctId, vAlert
    vAlertMsg = fIf(vAlert = 1, fPhraH(000947), fPhraH(000948))
  End If

  '...get max users, rights and alert needs
  sGetCust svCustId

  If Request.QueryString("vDelete").Count = 1 Then
    vMemb_No = Request.QueryString("vDelete")
    sDeleteMemb
    sDeleteEcomByMembNo
    Response.Redirect vNext

  ElseIf Request.Form("vHidden").Count = 1 Then
    sExtractMemb

    vPrograms = Replace(Request("vPrograms"), ",", "")
    If Instr(vMemb_Programs, vPrograms) = 0 Then
      vMemb_Programs = Trim(vMemb_Programs & " " & vPrograms)
    End If   
    vMemb_ProgramsLength = Cint(fDefault(Request("vMemb_ProgramsLength"), 0))
    If Len(vMemb_Programs) > vMemb_ProgramsLength Then vMemb_ProgramsAdded = Now()

    If svMembLevel = 3 Then 
      vMemb_LastAssignedBy = svMembNo
    End If
    
    If fNoValue(vMemb_Id) Then
      vMessage = fPhraH(001212)
    Else
      vMembId = Request("vMembId")
      If spMembExistsById (svCustAcctId, vMemb_Id) And (vMemb_No = 0 Or vMembId <> vMemb_Id) Then       
        vMessage = fPhraH(001213)
        If vMemb_Id <> vMembId Then vMemb_Id = vMembId '...put back original ID
      Else
        sAddMemb  svCustAcctId

        Response.Redirect fDefault(vNext, "Users.asp")
      End If
    End If 

  ElseIf Request.QueryString("vMembNo").Count = 1 Then
    vMemb_No = Request.QueryString("vMembNo")
    sGetMemb vMemb_No
    vMemb_ProgramsLength = Len(vMemb_Programs) '...get the length of this field so if it increases we can update the date

  Else  
    sGetMemb svMembNo

  End If

  If fNoValue(vMemb_Level) Then vMemb_Level = 2     

  bNoEditProfile = fIf(vMemb_No = svMembNo And svMembLevel < 4, True, False)

  vMemb_Active = fDefault(vMemb_Active, 1)

%>

<html>

<head>
  <meta charset="UTF-8">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>User Profiles</title>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <script>
    function validate(theForm) {

      if (theForm.vMemb_Id.value.length < 4 || theForm.vMemb_Id.value.length > 64 || theForm.vMemb_Id.value.match(rePassword)==null) {
        alert("Please enter a valid Password (4-64 chars).");
        theForm.vMemb_Id.focus();
        return (false);
      }

/*    
      if (theForm.vMemb_Email.value.length > 0 && theForm.vMemb_Email.value.match(reEmail)==null) { 
        alert("Please enter a valid Email Address.");
        theForm.vMemb_Email.focus();
        return (false);
      }
*/
      return (true);
    }
  </script>
</head>

<body>

  <% Server.Execute vShellHi %>

  <table width="100%">
    <tr>
      <td width="100%" valign="top" align="center">

        <h1>
          <% If vMemb_No = svMembNo Then %>
          <!--webbot bot='PurpleText' PREVIEW='My Profile'--><%=fPhra(000185)%>
          <% ElseIf vMemb_No = 0 Then %>
          <!--webbot bot='PurpleText' PREVIEW='Add a Learner'--><%=fPhra(000370)%>
          <% Else %>
          <!--webbot bot='PurpleText' PREVIEW='Learner Profile'--><%=fPhra(000371)%>
          <% End If %> 
        </h1>

        <% =fIf(Len(vMessage)>0,"<h5>" & vMessage & "</h5>", "") %>


        <% If bAlertOk Then %>
        <table border="1" cellspacing="0" cellpadding="5" bordercolor="#DDEEF9" style="border-collapse: collapse" width="80%">
          <tr>
            <td align="center">
              <%=fIf(Len(vAlertMsg) > 0, "<h5>" & vAlertMsg & "</h5>", "")%>

              <% If vCust_EcomG2alert Then %>

              <p align="left" class="red">
                <!--webbot bot='PurpleText' PREVIEW='This site is currently configured to automatically alert a Learner by email whenever program(s) are assigned to his/her profile. The email alert provides access instructions to the Learner and indicates that program(s) have been assigned by you, the account Facilitator. A Learner must have a valid email address in his/her profile below in order to qualify for email alerts. If this feature is not needed for any of your Learners, please click <b>Disable</b> below.'--><%=fPhra(001567)%>
              </p>
              <p align="center">
                <input onclick="location.href = 'UserGroup.asp?vAlert=0&amp;vMembNo=<%=vMemb_No%>'" type="button" value="<%=bDisable%>" name="bDisable" class="button">
              </p>

              <% If vCust_ParentId = "2962" Then  %>
              <p align="left">Are your Learners not receiving emails?  Please refer to page 18 of the Facilitator manual for details on how to proceed. If necessary, <a target="_blank" href="/gold/vuReporting/AccountTaskedit.aspx?AccountID=<%=svCustAcctId%>">click here to edit the email address from which your email alerts are sent</a> (defaults to using the Facilitator email address).&ensp;NOTE: In order for our system to send emails using your email address, your IT department may need to add the domain vubiz.com to your organization's SPF records.  An SPF - “Sender Policy Framework” - is setup so that specified external systems can send email on your behalf.  If your IT people need to speak to someone technical regarding this issue, please have them contact <a href="mailto:support@vubiz.com?subject=Email Alert Editor (<%=svCustId%>)">support@vubiz.com</a>.</p>
              <% End If %>

              <% Else %>
              <p align="left" class="red">
                <!--webbot bot='PurpleText' PREVIEW='This site is currently NOT configured to automatically email alert learners whenever you assign them content. If you would like this feature enabled click <b>Enable</b> below.'--><%=fPhra(001568)%>
                NOTE: Clicking Enable will activate the system for all Learners who are assigned program(s) moving forward. Learners will not receive email alerts on programs assigned in the past while the service was disabled.
              </p>
              <p align="center">
                <input onclick="location.href = 'UserGroup.asp?vAlert=1&amp;vMembNo=<%=vMemb_No%>'" type="button" value="<%=bEnable%>" name="bEnable" class="button">
              </p>
              <% End If %>
            </td>
          </tr>
        </table>
        <br>
        <% End If %>
      </td>
    </tr>
  </table>

  <form method="POST" action="UserGroup.asp" target="_self" onsubmit="return validate(this)">

    <input type="hidden" name="vHidden" value="Y">
    <input type="hidden" name="vMemb_No" value="<%=vMemb_No%>">
    <input type="hidden" name="vMembId" value="<%=vMemb_Id%>">
    <input type="hidden" name="vNext" value="<%=vNext%>">
    <input type="hidden" name="vMemb_Duration" value="<%=vMemb_Duration%>">
    <input type="hidden" name="vMemb_Manager" value="<%=fSqlBoolean(vMemb_Manager)%>">
    <input type="hidden" name="vMemb_ProgramsLength" value="<%=vMemb_ProgramsLength%>">

    <div align="center">

      <% If svMembLevel < 3 Or bNoEditProfile Then %>
      <tr>
        <th align="right" width="25%" valign="top"><%=fIf(svCustPwd, fPhraH(000411), fPhraH(000211))%> : </th>
        <td width="75%" valign="top" align="left"><%=vMemb_Id%></td>
      </tr>
      <input type="hidden" name="vMemb_Id" value="<%=vMemb_Id%>">
      <% Else %>
      <tr>
        <th align="right" width="25%" valign="top"><%=fIf(svCustPwd, fPhraH(000411), fPhraH(000211))%> : </th>
        <td width="75%" valign="top" align="left">
          <input type="text" size="20" name="vMemb_Id" value="<%=vMemb_Id%>" maxlength="64"><br>
          <!--webbot bot='PurpleText' PREVIEW='Must be unique using only English alpha, numeric and &quot;_.-@&quot; characters.'--><%=fPhra(000372)%>
        </td>
      </tr>
      <% End If %>

      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%>
          :</th>
        <td width="75%" valign="top" align="left">

          <% If bNoEditProfile Then %>
          <%=vMemb_FirstName%>
          <input type="hidden" name="vMemb_FirstName" value="<%=vMemb_FirstName%>">
          <% Else %>
          <input type="text" size="32" name="vMemb_FirstName" value="<%=vMemb_FirstName%>" maxlength="32">
          <% End If %>
        </td>
      </tr>

      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%>
          :</th>
        <td width="75%" valign="top" align="left">

          <% If bNoEditProfile Then %>
          <%=vMemb_LastName%>
          <input type="hidden" name="vMemb_LastName" value="<%=vMemb_LastName%>">
          <% Else %>
          <input type="text" size="32" name="vMemb_LastName" value="<%=vMemb_LastName%>" maxlength="64">
          <% End If %>
        </td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='Email Address'--><%=fPhra(000126)%>
          :</th>
        <td width="75%" valign="top" align="left">

          <% If bNoEditProfile Then %>
          <%=vMemb_Email%>
          <input type="hidden" name="vMemb_Email" value="<%=vMemb_Email%>">
          <% Else %>
          <input type="text" size="32" name="vMemb_Email" value="<%=vMemb_Email%>" maxlength="128">
          <% End If %>
        </td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='Organization'--><%=fPhra(000470)%>
          :</th>
        <td width="75%" valign="top" align="left">

          <% If bNoEditProfile Then %>
          <%=vMemb_Organization%>
          <input type="hidden" name="vMemb_Organization" value="<%=vMemb_Organization%>">
          <% Else %>
          <input type="text" size="46" name="vMemb_Organization" value="<%=vMemb_Organization%>" maxlength="128">
          <% End If %>
        </td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='Memo'--><%=fPhra(000173)%>
          :</th>
        <td width="75%" valign="top" align="left">
          <% If bNoEditProfile Then %>
          <%=vMemb_Memo%>
          <input type="hidden" name="vMemb_Memo" value="<%=vMemb_Memo%>">
          <% Else %>
          <input type="text" size="46" name="vMemb_Memo" value="<%=vMemb_Memo%>" maxlength="128">
          <% End If %>
        </td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='First Visit'--><%=fPhra(000157)%>
          : </th>
        <td width="75%" valign="top" align="left">

          <% If svMembLevel < 4 Then %>
          <%=fFormatDate(vMemb_FirstVisit)%>
          <input type="hidden" name="vMemb_FirstVisit" value="<%=fFormatDate(vMemb_FirstVisit)%>">
          <% Else %>
          <input type="text" name="vMemb_FirstVisit" size="20" value="<%=fFormatSqlDate (vMemb_FirstVisit)%>">
          ie <% =fFormatSqlDate(Now)%>.<br>
          <!--webbot bot='PurpleText' PREVIEW='Do not leave empty or it will revert to today&#39;s date.'--><%=fPhra(001569)%>
          <% End If %>
        </td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='Active'--><%=fPhra(000063)%>
          :</th>
        <td width="75%" valign="top" align="left">

          <% If bNoEditProfile Then %>
          <%=vMemb_Active%>
          <input type="hidden" name="vMemb_Active" value="<%=vMemb_Active%>">
          <% Else %>
          <input type="radio" name="vMemb_Active" value="0" <%=fcheck(0, fsqlboolean(vmemb_active))%>><!--webbot bot='PurpleText' PREVIEW='No'--><%=fPhra(000189)%>&nbsp;&nbsp; 
          <input type="radio" name="vMemb_Active" value="1" <%=fcheck(1, fsqlboolean(vmemb_active))%>><!--webbot bot='PurpleText' PREVIEW='Yes'--><%=fPhra(000024)%><br>
          <!--webbot bot='PurpleText' PREVIEW='Allows or disallows access to this service.'--><%=fPhra(000069)%>
          <% End If %>
        </td>
      </tr>


      <% If svMembLevel > 3 Then %>
      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='Learner Level'--><%=fPhra(000373)%>
          :</th>
        <td width="75%" valign="top" align="left">
          <input type="radio" name="vMemb_Level" value="2" <%=fcheck(2, vmemb_level)%>>2:<!--webbot bot='PurpleText' PREVIEW='Learner'--><%=fPhra(000165)%>
          (<!--webbot bot='PurpleText' PREVIEW='can access content and assessments'--><%=fPhra(000093)%>)<br>
          <input type="radio" name="vMemb_Level" value="3" <%=fcheck(3, vmemb_level)%>>3:<!--webbot bot='PurpleText' PREVIEW='Facilitator'--><%=fPhra(000139)%>
          (<!--webbot bot='PurpleText' PREVIEW='can add members and monitor progress'--><%=fPhra(000082)%>)<br>
          <% If svMembLevel > 3 Then %>
          <input type="radio" name="vMemb_Level" value="4" <%=fcheck(4, vmemb_level)%>>4: Manager (can access advanced features)<br>
          <% End If %> <% If svMembLevel > 4 Then %>
          <input type="radio" name="vMemb_Level" value="5" <%=fcheck(5, vmemb_level)%>>5: Administrator 
          <% End If %>
        </td>
      </tr>
      <% Else %>
      <input type="hidden" name="vMemb_Level" value="<%=vMemb_Level%>">
      <% End If %>


      <tr>
        <th align="right" width="25%" valign="top">&nbsp;</th>
        <td width="75%" valign="top" align="left">&nbsp;</td>
      </tr>

      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='Programs Assigned'--><%=fPhra(001570)%>
          :</th>
        <td width="75%" valign="top" align="left">

          <% If vMemb_No = svMembNo And svMembLevel = 3 Then %>
          <div style="border: 1px solid red; padding: 10px; text-align: center" class="red">
            <!--webbot bot='PurpleText' PREVIEW='DO NOT ASSIGN PROGRAMS TO YOUR OWN PROFILE !! <br>FACILITATORS GET ACCESS AUTOMATICALLY TO ALL PROGRAMS UNDER THE MY CONTENT TAB.'--><%=fPhra(001228)%>
          </div>
          <% End If %>

          <% If svMembLevel < 5 And Not svMembManager Then %>
          <input type="hidden" name="vMemb_Programs" value="<%=vMemb_Programs%>">
          <span style="background-color: #FFFF00"><%=fIf (Len(vMemb_Programs) = 0, "No program(s) currently assigned", vMemb_Programs)%></span>
          <% Else %>
          <input type="text" name="vMemb_Programs" value="<%=vMemb_Programs%>" style="width: 100%"><br>Only editable by administrators 
        <% End If %>

          <% 
          i = fEcomGroupProgs(vMemb_Programs)
          If i <> "" Then 
          %>
          <p>
            <!--webbot bot='PurpleText' PREVIEW='You need to assign one or more programs to this learner from the list of programs below. Click to highlight a program title, then click the Update button to apply your selection and save the learner profile. Use Ctrl+Click to make multiple course selections.'--><%=fPhra(001571)%>
            <font color="#FF0000">&nbsp;<br><br>
          <!--webbot bot='PurpleText' PREVIEW='Note'--><%=fPhra(001420)%>:
          <!--webbot bot='PurpleText' PREVIEW='Facilitators can assign one or more programs to a Learner, but once updated, these program(s) cannot be unassigned.'--><%=fPhra(000835)%>
          <!--webbot bot='PurpleText' PREVIEW='Facilitators must contact their Account Manager for assistance in the reassignment of programs.'--><%=fPhra(000836)%> </font>
          </p>
          <!-- This creates the dropdown and hidden field-->
          <%=i%>
          <% End If %>
        </td>
      </tr>

      <% If svMembLevel < 5 And Not svMembManager Then %>
      <input type="hidden" name="vMemb_Expires" value="<%=vMemb_Expires%>">
      <% Else %>
      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='Programs Expire'--><%=fPhra(000840)%>
          : </th>
        <td width="75%" valign="top" align="left">
          <input type="text" name="vMemb_Expires" size="11" value="<%=Trim(fFormatSqlDate (vMemb_Expires))%>">
          MMM DD, YYYY (ie: <% =fFormatSqlDate(Now + 90)%>)<br>
          <!--webbot bot='PurpleText' PREVIEW='If entered, signifies date that the above program(s) expire.'--><%=fPhra(000841)%>
        </td>
      </tr>
      <% End If %>

      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='Last Assigned By'--><%=fPhra(001595)%>
          : </th>
        <td width="75%" valign="top" align="left"><%=fDefault(fMembName (vMemb_LastAssignedBy), "N/A")%> - (<!--webbot bot='PurpleText' PREVIEW='Facilitator who last assigned content to this learner.'--><%=fPhra(001596)%>)</td>
      </tr>

      <% If bAlertOk And vCust_EcomG2alert Then %>
      <%  vMemb_EcomG2alert = fDefault(vMemb_EcomG2alert, 1) %>
      <tr>
        <th align="right" width="25%" valign="top">Above Programs Updated :</th>
        <td width="75%" valign="top" align="left"><%=fFormatSqlDate (vMemb_ProgramsAdded)%> </td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='Email Alert'--><%=fPhra(000127)%>
          ?</th>
        <td width="75%" valign="top" align="left">
          <input type="radio" name="vMemb_EcomG2Alert" value="1" <%=fcheck(fsqlboolean(vmemb_ecomg2alert), 1)%>><!--webbot bot='PurpleText' PREVIEW='Yes'--><%=fPhra(000024)%>&nbsp;
          <input type="radio" name="vMemb_EcomG2Alert" value="0" <%=fcheck(fsqlboolean(vmemb_ecomg2alert), 0)%>><!--webbot bot='PurpleText' PREVIEW='No'--><%=fPhra(000189)%><br>
          <!--webbot bot='PurpleText' PREVIEW='Turning this off (ie clicking <b>No</b>, will suspend the automatic email alert for this individual learner.&nbsp; When you turn it back an alert will be sent at the next scheduled run.'--><%=fPhra(000946)%>
        </td>
      </tr>
      <% End If %>









      <tr>
        <td align="center" width="100%" valign="top" colspan="2">

          <% If svCustAcctId <> fDefault(vMemb_AcctId, svCustAcctId) Then %>
          <h5><br><br><br>
            <!--webbot bot='PurpleText' PREVIEW='Learner Profiles accessed from another Account are Read Only.'--><%=fPhra(001265)%><br>
            <!--webbot bot='PurpleText' PREVIEW='You cannot Update or Delete this Learner&#39;s Profile'--><%=fPhra(001572)%>.<br></h5>

          <% Else %>

          <br>
          <div style="border: 1px solid red; margin: 20px 0px 20px 0px; padding: 10px;" class="red">

            <% If svMembManager Then  %>
              Note: Super Managers cannot modify their profile nor assign content.
            <% Else %>
            <!--webbot bot='PurpleText' PREVIEW='Update with caution.'--><%=fPhra(000838)%>
            &nbsp;
            <!--webbot bot='PurpleText' PREVIEW='Facilitators can assign one or more programs to a Learner, but once updated, these program(s) cannot be unassigned.'--><%=fPhra(000835)%>
            <% End If %>
          </div>

          <% If vMemb_No = svMembNo And svMembLevel = 3 Then %>
          <div style="border: 1px solid red; margin: 20px 0px 20px 0px; padding: 10px; text-align: center" class="red">
            <!--webbot bot='PurpleText' PREVIEW='DO NOT ASSIGN PROGRAMS TO YOUR OWN PROFILE !! <br>FACILITATORS GET ACCESS AUTOMATICALLY TO ALL PROGRAMS UNDER THE MY CONTENT TAB.'--><%=fPhra(001228)%>
          </div>
          <% End If %>

          <input onclick="location.href = '<%=vNext%>'" type="button" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button085">
          <%=f10%>

          <% If Not svMembManager Then %>
          <%=f10%>
          <input type="submit" value="<%=bUpdate%>" name="bUpdate" class="button085">
          <% End If %>

          <% If svMembLevel = 5 Then %>
          <%=f10%>
          <input onclick="jconfirm('UserGroup.asp?vDelete=<%=vMemb_No%>', '&lt<!--webbot bot='PurpleText' PREVIEW='gt;Ok to delete?&lt'--><%=fPhra(001573)%>gt;')" type="button" value="<%=bDelete%>" name="bDelete" class="button085">
          <% End If %>

          <% End If %>

          <br><br><br><a href="Users.asp?vNext=<%=vNext%>">
            <!--webbot bot='PurpleText' PREVIEW='Learner Report'--><%=fPhra(000367)%></a><%=f10%> <a href="UserGroup.asp?vMembNo=0">
              <!--webbot bot='PurpleText' PREVIEW='Add a Learner'--><%=fPhra(000370)%></a>

          <% If svMembLevel = 5 Then %>
          <%=f10%>
          <a <%=fstatx%> href="User.asp?vMembNo=<%=vMemb_No%>">Full Learner Profile</a>
          <% End If %>

          <br><br>
        </td>
      </tr>


















    </div>

  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


