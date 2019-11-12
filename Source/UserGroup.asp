<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<%
  Dim vMessage, vPrograms, vNext, vMemb_ProgramsLength, vAlert, bAlertOk, vAlertMsg, vMembId, bNoEditProfile
  
  vNext    = fDefault(Request("vNext"), "Users_O.asp?vSort=id&vStart=" & vMemb_Id)
  vMessage = Request.QueryString("vMessage")

  '...used in translation engine to Id Type
  p0 = fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")

  '...turn this on if the parent site offers email alerts to their G2 sites
  bAlertOk = fCustParentG2alertOk (svCustAcctId) 

  vAlert = Request("vAlert")
  If vAlert <> "" Then
    sUpdateCustG2alert svCustAcctId, vAlert
    vAlertMsg = fIf(vAlert = 1, "<!--{{-->Email Alert Service has been Enabled<!--}}-->", "<!--{{-->Email Alert Service has been Disabled<!--}}-->")
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
    
    If Not fMembIdOk (vMemb_Id)	Then
      vMessage = "<!--{{-->Unable to add/update record because the<!--}}-->" & " " & fIf(svCustPwd Or vCust_ChannelV8, "<!--{{-->Id<!--}}-->", "<!--{{-->Password<!--}}-->") & " " & "<!--{{-->contains an invalid character.<!--}}-->"
    ElseIf fNoValue(vMemb_Id) Then
      vMessage = "<!--{{-->Unable to update record without a unique ^0.<!--}}-->"
    Else
      vMembId = Request("vMembId")
      If spMembExistsById (svCustAcctId, vMemb_Id) And (vMemb_No = 0 Or vMembId <> vMemb_Id) Then       
        vMessage = "<!--{{-->That ^0 is already on file!<!--}}-->"
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
  <title>UserGroup</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <script>
    function validate(theForm) {

      var id = theForm.vMemb_Id.value.toUpperCase();
      if (id.length < 4 || id.length > 64 || id.match(rePassword)==null) {
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


  <h1><% If vMemb_No = svMembNo Then %><!--[[-->My Profile<!--]]--><% ElseIf vMemb_No = 0 Then %><!--[[-->Add a Learner<!--]]--><% Else %><!--[[-->Learner Profile<!--]]--><% End If %> </h1>

  <% =fIf(Len(vMessage)>0,"<h5 style='margin-bottom:20px;'>" & vMessage & "</h5>", "") %>


  <% If bAlertOk Then %>

    <div style="text-align:center; width:600px; margin: 0 auto 20px auto;">  

      <%=fIf(Len(vAlertMsg) > 0, "<h5>" & vAlertMsg & "</h5>", "")%>

      <% If vCust_EcomG2alert Then %>

        <p class="red" style="text-align:left;"><!--[[-->This site is currently configured to automatically alert a Learner by email whenever program(s) are assigned to his/her profile. The email alert provides access instructions to the Learner and indicates that program(s) have been assigned by you, the account Facilitator. A Learner must have a valid email address in his/her profile below in order to qualify for email alerts. If this feature is not needed for any of your Learners, please click <b>Disable</b> below.<!--]]--></p>
        <div style="text-align:center; margin:20px;"><input onclick="location.href = 'UserGroup.asp?vAlert=0&vMembNo=<%=vMemb_No%>'" type="button" value="<%=bDisable%>" name="bDisable" class="button"></div>

        <% If vCust_ParentId = "2962" Then  %>

        <p class="red" style="text-align:left;">
<!--      Are your Learners not receiving emails?  Please refer to page 18 of the Facilitator manual for details on how to proceed. If necessary, <a target="_blank" href="/gold/vuReporting/AccountTaskedit.aspx?AccountID=<%=svCustAcctId%>">click here to edit the email address from which your email alerts are sent</a> (defaults to using the Facilitator email address).&ensp;-->
          Are your Learners not receiving emails?  Please refer to page 18 of the Facilitator manual for details on how to proceed. If necessary and you require changes to be made to the default sender (Facilitator’s Email), <a target="_blank" href="/gold/vuReporting/AccountTaskedit.aspx?AccountID=<%=svCustAcctId%>">click here to edit the email address from which your email alerts are sent</a>.  Click Edit, insert support@vubiz.com in the Sender field, then click Update.&ensp;        
          NOTE: In order for our system to send emails using your email address, your internal IT department may need to add *.vubiz.com (or IPs 104.45.154.149 & 104.41.147.19 & 191.237.26.225) to your whitelist and to your SPF records. An SPF - 'Sender Policy Framework' - is setup so that specified external systems can send email on your behalf. If your IT people need to speak to someone technical regarding this issue, please have them contact <a href="mailto:support@vubiz.com?subject=Email Alert Editor (<%=svCustId%>)">support@vubiz.com</a>.
        </p>

        <% End If %>

      <% Else %>

<!--      <div class="red" style="text-align:left; margin:20px; padding:20px; border:1px solid red;">-->
          <div style="border: 1px solid red; width:600px; margin: 20px auto;  padding: 10px; text-align:left;" class="red">
          <!--[[-->This site is currently NOT configured to automatically email alert learners whenever you assign them content. If you would like this feature enabled click <b>Enable</b> below. NOTE: Clicking Enable will activate the system for all Learners who are assigned program(s) moving forward. Learners will not receive email alerts on programs assigned in the past while the service was disabled.<!--]]-->
          <div style="text-align:center; margin-top:20px;"><input onclick="location.href = 'UserGroup.asp?vAlert=1&amp;vMembNo=<%=vMemb_No%>'" type="button" value="<%=bEnable%>" name="bEnable" class="button"></div>
        </div>

      <% End If %> 

     </div>

  <% End If %> 


  <form method="POST" action="UserGroup.asp" target="_self" onsubmit="return validate(this)">

    <input type="hidden" name="vHidden" value="Y">
    <input type="hidden" name="vMemb_No" value="<%=vMemb_No%>">
    <input type="hidden" name="vMembId" value="<%=vMemb_Id%>">
    <input type="hidden" name="vNext" value="<%=vNext%>">
    <input type="hidden" name="vMemb_Duration" value="<%=vMemb_Duration%>">
    <input type="hidden" name="vMemb_Manager" value="<%=fSqlBoolean(vMemb_Manager)%>">
    <input type="hidden" name="vMemb_ProgramsLength" value="<%=vMemb_ProgramsLength%>">


      <table id="UserGroup">

        <% If svMembLevel < 3 Or bNoEditProfile Then %>
        <tr>
          <th><%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%> : </th>
          <td><%=vMemb_Id%></td>
        </tr>
        <input type="hidden" name="vMemb_Id" value="<%=vMemb_Id%>">
        <% Else %>
        <tr>
          <th><%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%> : </th>
          <td>
            <input type="text" size="20" name="vMemb_Id" value="<%=vMemb_Id%>" maxlength="64"><br>
            <!--[[-->Must be unique using only English alpha, numeric and !@$%^*()_-{}[];,.: characters.&nbsp;Value is NOT case sensitive.<!--]]-->
          </td>
        </tr>
        <% End If %>

        <tr>
          <th>
            <!--[[-->First Name<!--]]--> :</th>
          <td>

            <% If bNoEditProfile Then %>
            <%=vMemb_FirstName%>
            <input type="hidden" name="vMemb_FirstName" value="<%=vMemb_FirstName%>">
            <% Else %>
            <input type="text" size="32" name="vMemb_FirstName" value="<%=vMemb_FirstName%>" maxlength="32">
            <% End If %>

          </td>
        </tr>
        <tr>
          <th>
            <!--[[-->Last Name<!--]]-->:</th>
          <td>

            <% If bNoEditProfile Then %>
            <%=vMemb_LastName%>
            <input type="hidden" name="vMemb_LastName" value="<%=vMemb_LastName%>">
            <% Else %>
            <input type="text" size="32" name="vMemb_LastName" value="<%=vMemb_LastName%>" maxlength="64">
            <% End If %>

          </td>
        </tr>
        <tr>
          <th>
            <!--[[-->Email Address<!--]]--> :</th>
          <td>

            <% If bNoEditProfile Then %>
            <%=vMemb_Email%>
            <input type="hidden" name="vMemb_Email" value="<%=vMemb_Email%>">
            <% Else %>
            <input type="text" size="32" name="vMemb_Email" value="<%=vMemb_Email%>" maxlength="128">
            <% End If %>

          </td>
        </tr>
        <tr>
          <th>
            <!--[[-->Organization<!--]]--> :</th>
          <td>

            <% If bNoEditProfile Then %>
            <%=vMemb_Organization%>
            <input type="hidden" name="vMemb_Organization" value="<%=vMemb_Organization%>">
            <% Else %>
            <input type="text" size="46" name="vMemb_Organization" value="<%=vMemb_Organization%>" maxlength="128">
            <% End If %>

          </td>
        </tr>
        <tr>
          <th>
            <!--[[-->Memo<!--]]--> :</th>
          <td>
            <% If bNoEditProfile Then %>
            <%=vMemb_Memo%>
            <input type="hidden" name="vMemb_Memo" value="<%=vMemb_Memo%>">
            <% Else %>
            <input type="text" size="46" name="vMemb_Memo" value="<%=vMemb_Memo%>" maxlength="128">
            <% End If %>
          </td>
        </tr>
        <tr>
          <th>
            <!--[[-->First Visit<!--]]--> : </th>
          <td>

            <% If svMembLevel < 4 Then %>
            <%=fFormatDate(vMemb_FirstVisit)%>
            <input type="hidden" name="vMemb_FirstVisit" value="<%=fFormatDate(vMemb_FirstVisit)%>">
            <% Else %>
            <input type="text" name="vMemb_FirstVisit" size="20" value="<%=fFormatSqlDate (vMemb_FirstVisit)%>">
            ie <% =fFormatSqlDate(Now)%>.<br>
            <!--[[-->Do not leave empty or it will revert to today&#39;s date.<!--]]-->
            <% End If %>

          </td>
        </tr>
        <tr>
          <th>
            <!--[[-->Active<!--]]--> :</th>
          <td>

            <% If bNoEditProfile Then %>
            <%=vMemb_Active%>
            <input type="hidden" name="vMemb_Active" value="<%=vMemb_Active%>">
            <% Else %>
            <input type="radio" name="vMemb_Active" value="0" <%=fcheck(0, fsqlboolean(vmemb_active))%>><!--[[-->No<!--]]-->&nbsp;&nbsp; 
          <input type="radio" name="vMemb_Active" value="1" <%=fcheck(1, fsqlboolean(vmemb_active))%>><!--[[-->Yes<!--]]--><br>
            <!--[[-->Allows or disallows access to this service.<!--]]-->
            <% End If %>

          </td>
        </tr>


        <% If svMembLevel > 3 Then %>
        <tr>
          <th><!--[[-->Learner/Security Level<!--]]--> :</th>
          <td>
            <input type="radio" name="vMemb_Level" value="2" <%=fcheck(2, vmemb_level)%>>2:<!--[[-->Learner<!--]]-->
            (<!--[[-->can access content and assessments<!--]]-->)<br>
            <input type="radio" name="vMemb_Level" value="3" <%=fcheck(3, vmemb_level)%>>3:<!--[[-->Facilitator<!--]]-->
            (<!--[[-->can add members and monitor progress<!--]]-->)<br>
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



        <% '...Group 3 for Advanced Channels
           If Not vCust_ChannelReportsTo Or svMembLevel < 4 Or vMemb_Level > 2 Then        
        %>
        <input type="hidden" name="vMemb_Group3" value="<%=vMemb_Group3%>">
        <%
           End If

           If vCust_ChannelReportsTo Then
        %>
        <tr>
          <th class="notice">Reports To :</th>
          <td class="notice">
            <% If vMemb_Level > 2 Then %>
            Reports To values can only be assigned to Learners.
            <% Else              
                 If svMembLevel > 3 Then %>
            <select size="6" name="vMemb_Group3" class="c2">
              <option value="0" selected>Unassigned...</option>
              <% =fMembFacsDropdown (vMemb_Group3) %>
            </select><br />
            The selected facilitator will be able to report on this learner.&ensp;Note: "Unassigned" learners can be accessed by any facilitator.<br />[coming]
               <%
                 Else 
                   Response.Write fIf(fMembName (vMemb_Group3) = "", "Unassigned", fMembName (vMemb_Group3))
                 End If 
                
               End If 
               %>
          </td>
        </tr>
        <%
        End If
        %>

        <tr>
          <th>&nbsp;</th>
          <td>&nbsp;</td>
        </tr>

        <tr>
          <th>
            <!--[[-->Programs Assigned<!--]]--> :</th>
          <td>

            <% If vMemb_No = svMembNo And svMembLevel = 3 Then %>
            <div style="border: 1px solid red; padding: 10px; text-align: center" class="red">
              <!--[[-->DO NOT ASSIGN PROGRAMS TO YOUR OWN PROFILE !! <br>FACILITATORS GET ACCESS AUTOMATICALLY TO ALL PROGRAMS UNDER THE MY CONTENT TAB.<!--]]-->
            </div>
            <% End If %>

            <% If svMembLevel < 5 And Not svMembManager Then %>
            <input type="hidden" name="vMemb_Programs" value="<%=vMemb_Programs%>">
            <span style="background-color: #FFFF00"><%=fIf (Len(vMemb_Programs) = 0, "<!--{{-->No programs currently assigned<!--}}-->", vMemb_Programs)%></span>
            <% Else %>
            <input type="text" name="vMemb_Programs" value="<%=vMemb_Programs%>" style="width: 100%"><br>Only editable by administrators. Separate programs with a single space. Do not embed a comma. 
            <% End If %>

            <% 
              stop
          i = fEcomGroupProgs(vMemb_Programs)
          If i <> "" Then 
            %>
            <p>
              <!--[[-->You need to assign one or more programs to this learner from the list of programs below. Click to highlight a program title, then click the Update button to apply your selection and save the learner profile. Use Ctrl+Click to make multiple course selections.<!--]]-->
              <span style="color:#FF0000">&nbsp;<br><br>
                <!--[[-->Note<!--]]-->:
                <!--[[-->Facilitators can assign one or more programs to a Learner, but once updated, these program(s) cannot be unassigned.<!--]]-->
                <!--[[-->Facilitators must contact their Account Manager for assistance in the reassignment of programs.<!--]]--> 
              </span>
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
          <th><!--[[-->Programs Expire<!--]]--> :</th>
          <td>
            <input type="text" name="vMemb_Expires" size="11" value="<%=Trim(fFormatSqlDate (vMemb_Expires))%>">
            MMM DD, YYYY (ie: <% =fFormatSqlDate(Now + 90)%>)<br>
            <!--[[-->If entered, signifies date that the above program(s) expire.<!--]]-->
          </td>
        </tr>
        <% End If %>

        <tr>
          <th>
            <!--[[-->Last Assigned By<!--]]--> :</th>
          <td><%=fDefault(fMembName (vMemb_LastAssignedBy), "N/A")%> - (<!--[[-->Facilitator who last assigned content to this learner.<!--]]-->)</td>
        </tr>

        <% If bAlertOk And vCust_EcomG2alert Then %>
        <%  vMemb_EcomG2alert = fDefault(vMemb_EcomG2alert, 1) %>
        <tr>
          <th><!--[[-->Above Programs Updated<!--]]--> :</th>
          <td><%=fFormatSqlDate (vMemb_ProgramsAdded)%> </td>
        </tr>
        <tr>
          <th><!--[[-->Email Alert<!--]]--> ?</th>
          <td>
            <input type="radio" name="vMemb_EcomG2Alert" value="1" <%=fcheck(fsqlboolean(vmemb_ecomg2alert), 1)%>><!--[[-->Yes<!--]]-->&nbsp;
            <input type="radio" name="vMemb_EcomG2Alert" value="0" <%=fcheck(fsqlboolean(vmemb_ecomg2alert), 0)%>><!--[[-->No<!--]]--><br>
            <!--[[-->Turning this off (ie clicking <b>No</b>, will suspend the automatic email alert for this individual learner.&nbsp; When you turn it back an alert will be sent at the next scheduled run.<!--]]--></td>
        </tr>
        <% End If %>

        <tr>
          <td style="text-align:center;"" colspan="2">

            <% If svCustAcctId <> fDefault(vMemb_AcctId, svCustAcctId) Then %>
            <h5><br><br><br><!--[[-->Learner Profiles accessed from another Account are Read Only.<!--]]--><br><!--[[-->You cannot Update or Delete this Learner&#39;s Profile<!--]]-->.<br></h5>
            <% Else %>

            <div style="border: 1px solid red; width:600px; margin: 20px auto;  padding: 10px; text-align:left;" class="red">
              <% If vMemb_No = svMembNo And svMembManager Then  %>
              Note: Super Managers cannot modify their own profile.
              <% Else %>
              <!--[[-->Update with caution.<!--]]-->&nbsp;<!--[[-->You can assign one or more programs to a Learner, but once updated, these program(s) cannot be unassigned.<!--]]-->
              <% End If %>
            </div>

            <% If vMemb_No = svMembNo And svMembLevel = 3 Then %>
            <div style="border: 1px solid red; width:600px; margin: 20px auto;  padding: 10px; text-align:left;" class="red">
              <!--[[-->DO NOT ASSIGN PROGRAMS TO YOUR OWN PROFILE !! <br>FACILITATORS GET ACCESS AUTOMATICALLY TO ALL PROGRAMS UNDER THE MY CONTENT TAB.<!--]]-->
            </div>
            <% End If %>

            <input onclick="location.href = '<%=vNext%>'" type="button" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button085">

            <% If Not (vMemb_No = svMembNo And svMembManager) Then %>
            <%=f10%>
            <input type="submit" value="<%=bUpdate%>" name="bUpdate" class="button085">
            <% End If %>

            <% If (svMembLevel = 5 Or svMembManager) And (vMemb_No <> svMembNo) Then %>
            <%=f10%>
            <input onclick="jconfirm('UserGroup.asp?vDelete=<%=vMemb_No%>', '<!--[[-->Ok to delete?<!--]]-->')" type="button" value="<%=bDelete%>" name="bDelete" class="button085">
            <% End If %>

            <% End If %>

            <br><br><br>
            <a href="Users.asp?vNext=<%=vNext%>"><!--[[-->Learner Report<!--]]--></a><%=f10%> 
            <a href="UserGroup.asp?vMembNo=0"><!--[[-->Add a Learner<!--]]--></a>
            <% If svMembLevel = 5 Then %><%=f10%>
            <a <%=fstatx%> href="User.asp?vMembNo=<%=vMemb_No%>">Full Learner Profile</a>
            <% End If %>
            <br><br></td>

        </tr>

      </table>


  </form>

  <style>
    #UserGroup tr th { width: 25%; padding: 2px; }
    #UserGroup tr td { width: 75%; padding: 2px; }
  </style>


  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
