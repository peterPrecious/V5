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
    
    If Not fMembIdOk (vMemb_Id)	Then
      vMessage = fPhraH(001785) & " " & fIf(svCustPwd Or vCust_ChannelV8, fPhraH(000374), fPhraH(000211)) & " " & fPhraH(001786)
    ElseIf fNoValue(vMemb_Id) Then
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
  <title>UserGroup</title>
  <meta charset="UTF-8">
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


  <h1><% If vMemb_No = svMembNo Then %><!--webbot bot='PurpleText' PREVIEW='My Profile'--><%=fPhra(000185)%><% ElseIf vMemb_No = 0 Then %><!--webbot bot='PurpleText' PREVIEW='Add a Learner'--><%=fPhra(000370)%><% Else %><!--webbot bot='PurpleText' PREVIEW='Learner Profile'--><%=fPhra(000371)%><% End If %> </h1>

  <% =fIf(Len(vMessage)>0,"<h5 style='margin-bottom:20px;'>" & vMessage & "</h5>", "") %>


  <% If bAlertOk Then %>

    <div style="text-align:center; width:600px; margin: 0 auto 20px auto;">  

      <%=fIf(Len(vAlertMsg) > 0, "<h5>" & vAlertMsg & "</h5>", "")%>

      <% If vCust_EcomG2alert Then %>

        <p class="red" style="text-align:left;"><!--webbot bot='PurpleText' PREVIEW='This site is currently configured to automatically alert a Learner by email whenever program(s) are assigned to his/her profile. The email alert provides access instructions to the Learner and indicates that program(s) have been assigned by you, the account Facilitator. A Learner must have a valid email address in his/her profile below in order to qualify for email alerts. If this feature is not needed for any of your Learners, please click <b>Disable</b> below.'--><%=fPhra(001567)%></p>
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
          <!--webbot bot='PurpleText' PREVIEW='This site is currently NOT configured to automatically email alert learners whenever you assign them content. If you would like this feature enabled click <b>Enable</b> below. NOTE: Clicking Enable will activate the system for all Learners who are assigned program(s) moving forward. Learners will not receive email alerts on programs assigned in the past while the service was disabled.'--><%=fPhra(001719)%>
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
          <th><%=fIf(svCustPwd, fPhraH(000411), fPhraH(000211))%> : </th>
          <td><%=vMemb_Id%></td>
        </tr>
        <input type="hidden" name="vMemb_Id" value="<%=vMemb_Id%>">
        <% Else %>
        <tr>
          <th><%=fIf(svCustPwd, fPhraH(000411), fPhraH(000211))%> : </th>
          <td>
            <input type="text" size="20" name="vMemb_Id" value="<%=vMemb_Id%>" maxlength="64"><br>
            <!--webbot bot='PurpleText' PREVIEW='Must be unique using only English alpha, numeric and !@$%^*()_-{}[];,.: characters.&nbsp;Value is NOT case sensitive.'--><%=fPhra(001784)%>
          </td>
        </tr>
        <% End If %>

        <tr>
          <th>
            <!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%> :</th>
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
            <!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%>:</th>
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
            <!--webbot bot='PurpleText' PREVIEW='Email Address'--><%=fPhra(000126)%> :</th>
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
            <!--webbot bot='PurpleText' PREVIEW='Organization'--><%=fPhra(000470)%> :</th>
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
            <!--webbot bot='PurpleText' PREVIEW='Memo'--><%=fPhra(000173)%> :</th>
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
            <!--webbot bot='PurpleText' PREVIEW='First Visit'--><%=fPhra(000157)%> : </th>
          <td>

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
          <th>
            <!--webbot bot='PurpleText' PREVIEW='Active'--><%=fPhra(000063)%> :</th>
          <td>

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
          <th><!--webbot bot='PurpleText' PREVIEW='Learner/Security Level'--><%=fPhra(001611)%> :</th>
          <td>
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
            <!--webbot bot='PurpleText' PREVIEW='Programs Assigned'--><%=fPhra(001570)%> :</th>
          <td>

            <% If vMemb_No = svMembNo And svMembLevel = 3 Then %>
            <div style="border: 1px solid red; padding: 10px; text-align: center" class="red">
              <!--webbot bot='PurpleText' PREVIEW='DO NOT ASSIGN PROGRAMS TO YOUR OWN PROFILE !! <br>FACILITATORS GET ACCESS AUTOMATICALLY TO ALL PROGRAMS UNDER THE MY CONTENT TAB.'--><%=fPhra(001228)%>
            </div>
            <% End If %>

            <% If svMembLevel < 5 And Not svMembManager Then %>
            <input type="hidden" name="vMemb_Programs" value="<%=vMemb_Programs%>">
            <span style="background-color: #FFFF00"><%=fIf (Len(vMemb_Programs) = 0, fPhraH(001718), vMemb_Programs)%></span>
            <% Else %>
            <input type="text" name="vMemb_Programs" value="<%=vMemb_Programs%>" style="width: 100%"><br>Only editable by administrators. Separate programs with a single space. Do not embed a comma. 
            <% End If %>

            <% 
              stop
          i = fEcomGroupProgs(vMemb_Programs)
          If i <> "" Then 
            %>
            <p>
              <!--webbot bot='PurpleText' PREVIEW='You need to assign one or more programs to this learner from the list of programs below. Click to highlight a program title, then click the Update button to apply your selection and save the learner profile. Use Ctrl+Click to make multiple course selections.'--><%=fPhra(001571)%>
              <span style="color:#FF0000">&nbsp;<br><br>
                <!--webbot bot='PurpleText' PREVIEW='Note'--><%=fPhra(001420)%>:
                <!--webbot bot='PurpleText' PREVIEW='Facilitators can assign one or more programs to a Learner, but once updated, these program(s) cannot be unassigned.'--><%=fPhra(000835)%>
                <!--webbot bot='PurpleText' PREVIEW='Facilitators must contact their Account Manager for assistance in the reassignment of programs.'--><%=fPhra(000836)%> 
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
          <th><!--webbot bot='PurpleText' PREVIEW='Programs Expire'--><%=fPhra(000840)%> :</th>
          <td>
            <input type="text" name="vMemb_Expires" size="11" value="<%=Trim(fFormatSqlDate (vMemb_Expires))%>">
            MMM DD, YYYY (ie: <% =fFormatSqlDate(Now + 90)%>)<br>
            <!--webbot bot='PurpleText' PREVIEW='If entered, signifies date that the above program(s) expire.'--><%=fPhra(000841)%>
          </td>
        </tr>
        <% End If %>

        <tr>
          <th>
            <!--webbot bot='PurpleText' PREVIEW='Last Assigned By'--><%=fPhra(001595)%> :</th>
          <td><%=fDefault(fMembName (vMemb_LastAssignedBy), "N/A")%> - (<!--webbot bot='PurpleText' PREVIEW='Facilitator who last assigned content to this learner.'--><%=fPhra(001596)%>)</td>
        </tr>

        <% If bAlertOk And vCust_EcomG2alert Then %>
        <%  vMemb_EcomG2alert = fDefault(vMemb_EcomG2alert, 1) %>
        <tr>
          <th><!--webbot bot='PurpleText' PREVIEW='Above Programs Updated'--><%=fPhra(001716)%> :</th>
          <td><%=fFormatSqlDate (vMemb_ProgramsAdded)%> </td>
        </tr>
        <tr>
          <th><!--webbot bot='PurpleText' PREVIEW='Email Alert'--><%=fPhra(000127)%> ?</th>
          <td>
            <input type="radio" name="vMemb_EcomG2Alert" value="1" <%=fcheck(fsqlboolean(vmemb_ecomg2alert), 1)%>><!--webbot bot='PurpleText' PREVIEW='Yes'--><%=fPhra(000024)%>&nbsp;
            <input type="radio" name="vMemb_EcomG2Alert" value="0" <%=fcheck(fsqlboolean(vmemb_ecomg2alert), 0)%>><!--webbot bot='PurpleText' PREVIEW='No'--><%=fPhra(000189)%><br>
            <!--webbot bot='PurpleText' PREVIEW='Turning this off (ie clicking <b>No</b>, will suspend the automatic email alert for this individual learner.&nbsp; When you turn it back an alert will be sent at the next scheduled run.'--><%=fPhra(000946)%></td>
        </tr>
        <% End If %>

        <tr>
          <td style="text-align:center;"" colspan="2">

            <% If svCustAcctId <> fDefault(vMemb_AcctId, svCustAcctId) Then %>
            <h5><br><br><br><!--webbot bot='PurpleText' PREVIEW='Learner Profiles accessed from another Account are Read Only.'--><%=fPhra(001265)%><br><!--webbot bot='PurpleText' PREVIEW='You cannot Update or Delete this Learner&#39;s Profile'--><%=fPhra(001572)%>.<br></h5>
            <% Else %>

            <div style="border: 1px solid red; width:600px; margin: 20px auto;  padding: 10px; text-align:left;" class="red">
              <% If vMemb_No = svMembNo And svMembManager Then  %>
              Note: Super Managers cannot modify their own profile.
              <% Else %>
              <!--webbot bot='PurpleText' PREVIEW='Update with caution.'--><%=fPhra(000838)%>&nbsp;<!--webbot bot='PurpleText' PREVIEW='You can assign one or more programs to a Learner, but once updated, these program(s) cannot be unassigned.'--><%=fPhra(001597)%>
              <% End If %>
            </div>

            <% If vMemb_No = svMembNo And svMembLevel = 3 Then %>
            <div style="border: 1px solid red; width:600px; margin: 20px auto;  padding: 10px; text-align:left;" class="red">
              <!--webbot bot='PurpleText' PREVIEW='DO NOT ASSIGN PROGRAMS TO YOUR OWN PROFILE !! <br>FACILITATORS GET ACCESS AUTOMATICALLY TO ALL PROGRAMS UNDER THE MY CONTENT TAB.'--><%=fPhra(001228)%>
            </div>
            <% End If %>

            <input onclick="location.href = '<%=vNext%>'" type="button" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button085">

            <% If Not (vMemb_No = svMembNo And svMembManager) Then %>
            <%=f10%>
            <input type="submit" value="<%=bUpdate%>" name="bUpdate" class="button085">
            <% End If %>

            <% If (svMembLevel = 5 Or svMembManager) And (vMemb_No <> svMembNo) Then %>
            <%=f10%>
            <input onclick="jconfirm('UserGroup.asp?vDelete=<%=vMemb_No%>', '<!--webbot bot='PurpleText' PREVIEW='Ok to delete?'--><%=fPhra(000199)%>')" type="button" value="<%=bDelete%>" name="bDelete" class="button085">
            <% End If %>

            <% End If %>

            <br><br><br>
            <a href="Users.asp?vNext=<%=vNext%>"><!--webbot bot='PurpleText' PREVIEW='Learner Report'--><%=fPhra(000367)%></a><%=f10%> 
            <a href="UserGroup.asp?vMembNo=0"><!--webbot bot='PurpleText' PREVIEW='Add a Learner'--><%=fPhra(000370)%></a>
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


