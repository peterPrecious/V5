<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_TskH.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<%
  Dim vMessage, vNext, vMembId

  vNext    = Request("vNext")
  vMessage = Request.QueryString("vMessage")

  '...used in translation engine to Id Type
  p0 = fIf(svCustPwd, fPhraH(000411), fPhraH(000211))

  '...get max users and authoring rights
  sGetCust svCustId
  vMemb_AcctId = vCust_AcctId

  If Request.QueryString("vDelete").Count = 1 Then
    vMemb_No = Request.QueryString("vDelete")
    sDeleteMemb
    sDeleteEcomByMembNo
    Response.Redirect fDefault(vNext, "Users.asp")

  ElseIf Request.Form("vHidden").Count = 1 Then

    sExtractMemb

    If Not fMembIdOk (vMemb_Id)	Then
      vMessage = fPhraH(001785) & " " & fIf(svCustPwd Or vCust_ChannelV8, fPhraH(000374), fPhraH(000211)) & " " & fPhraH(001786)
    ElseIf fNoValue(vMemb_Id) Then
      vMessage = fPhraH(001212)
    Else
      vMembId = Ucase(Request("vMembId")) '...this is the value that was originally entered      
      '...ensure we are not adding a name that already exists
      If spMembExistsById (vMemb_AcctId, vMemb_Id) And (vMemb_No = 0 Or vMembId <> vMemb_Id) Then       
        vMessage = fPhraH(001213)
        If vMemb_Id <> vMembId Then vMemb_Id = vMembId '...put back original ID
      Else
        sAddMemb vMemb_AcctId
        Response.Redirect fDefault(vNext, "Users.asp")
      End If
    End If   

  ElseIf Request.QueryString("vMembNo").Count = 1 Then
    vMemb_No = Request.QueryString("vMembNo")
    If vMemb_No > 0 Then
      sGetMemb vMemb_No
    Else
      vMemb_Sponsor = 0
    End If

  Else  
    sGetMemb svMembNo

  End If

  If fNoValue(vMemb_Level) Then vMemb_Level = 2     
%>

<html>

<head>

  <title>User</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <script>
    function validate(theForm) {
  
      var id = theForm.vMemb_Id.value.toUpperCase();
      if (id.length < 4 || id.length > 64 || id.match(rePassword)==null) {
        if (theForm.vPassword.value == "check" ) {
          alert("Please enter a valid Learner Id (4-64 chars).");
        } else {
          alert("Please enter a valid Password (4-64 chars).");
        }
        theForm.vMemb_Id.focus();
        return (false);
      }    

      //  check password if used by this account and is a member (ie not fac/mgr)
      if (theForm.vPassword.value == "check" && theForm.vMemb_Level.value == 2) {
        var pwd = theForm.vMemb_Pwd.value.toUpperCase();
        if (pwd.length < 4 || pwd.value.length > 64 || pwd.match(rePassword)==null) {
          alert("Please enter a valid Password (4-64 chars).");
        theForm.vMemb_Pwd.focus();
        return (false);
        }
      }

/*
      if (theForm.vMemb_Email.value.length > 0 && theForm.vMemb_Email.value.match(reEmail)==null) { 
        alert("Please enter a valid Email Address.");
        theForm.vMemb_Email.focus();
        return (false);
      }

*/    
      if (theForm.vMemb_Duration.value.length > 0 && theForm.vMemb_Duration.value.match(reNumeric)==null) {
        alert("Please enter a valid Duration period (1-365)");
        theForm.vMemb_Duration.focus();
        return (false);
      }

      if (theForm.vMemb_Criteria != undefined) {
        if (theForm.vMemb_Criteria.selectedIndex < 0) {
          alert("Please select a Group.");
          theForm.vMemb_Criteria.focus();
          return (false);
        }            
      }

      return (true);
    }
  </script>
</head>

<body>

  <% Server.Execute vShellHi %>

  <form name="FrontPage_Form1" method="POST" action="User<%=fGroup%>.asp" target="_self" onsubmit="return validate(this)">

    <input type="hidden" name="vHidden" value="Y">
    <input type="hidden" name="vNext" value="<%=vNext%>">
    <input type="hidden" name="vMemb_No" value="<%=vMemb_No%>">
    <input type="hidden" name="vMembId" value="<%=vMemb_Id%>">

    <!-- this is the ID before update - incase we try to change it to an existing ID -->
    <table class="table">
      <tr>
        <td colspan="2" style="text-align: center; height: 50px;">
          <h1><!--webbot bot='PurpleText' PREVIEW='Add a Learner'--><%=fPhra(000370)%>/<!--webbot bot='PurpleText' PREVIEW='Learner Profile'--><%=fPhra(000371)%></h1>
          <% 
          If vCust_MaxUsers > 0 And (vCust_Level < 3 Or svMembLevel = 5 Or svMembManager) Then 
            p0 = vCust_MaxUsers
            p1 = fAllMembCount - 1
          %>
          <h6 align="left"><!--webbot bot='PurpleText' PREVIEW='Note: This account is limited to ^0 active or inactive learners.&nbsp; Once you reach the maximum you will be unable to add new learners or edit existing learners.&nbsp; You currently have ^1 learners on file.'--><%=fPhra(000512)%></h6>
          <% 
          End If 
  
          If Not fNoValue(vMessage) Then
          %>
          <h5 style="margin-bottom:20px;"><%=vMessage%></h5>
          <% 
          End If 
          %>
        </td>
      </tr>
      <tr>
        <th><%=fIf(svCustPwd or vCust_ChannelV8, fPhraH(000374), fPhraH(000211))%> :</th>
        <td>
          <% 
            If svMembLevel < 3 Then 
          %>
          <%=fIf(InStr(vMemb_Id, vPasswordx) > 0 , "********", vMemb_Id)%> <%=vMemb_Id%> <% = fIf (svMembLevel = 5, f10 & "(Learner No : " & vMemb_No & ")", "") %>
          <input type="hidden" name="vMemb_Id" value="<%=vMemb_Id%>">
          <% 
            Else 
          %>
          <input type="text" size="30" name="vMemb_Id" value="<%=vMemb_Id%>" maxlength="64">
          <% = fIf (svMembLevel = 5, f10 & "(Learner No : " & vMemb_No & ")", "") %> <br>
          <!--webbot bot='PurpleText' PREVIEW='Must be unique using only English alpha, numeric and !@$%^*()_-{}[];,.: characters.&nbsp;Value is NOT case sensitive.'--><%=fPhra(001784)%>
          <% 
            End If 
          %> 
        </td>
      </tr>
      <% 
        If vCust_Pwd or vCust_ChannelV8 Then 
      %>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Password'--><%=fPhra(000211)%> : </th>
        <td>
          <input type="text" size="30" name="vMemb_Pwd" value="<%=vMemb_Pwd%>" maxlength="64"><br>
          <!--webbot bot='PurpleText' PREVIEW='Use English alpha, numeric and !@$%^*()_-{}[];<>,.: characters.&nbsp;Value is NOT case sensitive.'--><%=fPhra(001782)%>
          <br><span style="color:red">Do not assign for facilitators or managers.</span>
        </td>
      </tr>
      <input type="hidden" name="vPassword" value="check">
      <% 
        Else 
      %>
      <input type="hidden" name="vPassword" value="ignore">
      <% 
        End If 
      %>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%>: </th>
        <td>
          <input type="text" size="30" name="vMemb_FirstName" value="<%=vMemb_FirstName%>" maxlength="32"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%>: </th>
        <td><input type="text" size="30" name="vMemb_LastName" value="<%=vMemb_LastName%>" maxlength="64"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Email Address'--><%=fPhra(000126)%>: </th>
        <td><input type="text" size="46" name="vMemb_Email" value="<%=vMemb_Email%>" maxlength="128"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Organization'--><%=fPhra(000470)%>: </th>
        <td>
          <input type="text" size="46" name="vMemb_Organization" value="<%=vMemb_Organization%>" maxlength="128"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Active'--><%=fPhra(000063)%>:</th>
        <td>
          <% 
          If fNoValue(vMemb_Active) Then vMemb_Active = 1 
          %>
          <input type="radio" name="vMemb_Active" value="0" <%=fcheck(0, fsqlboolean(vmemb_active))%>><!--webbot bot='PurpleText' PREVIEW='No'--><%=fPhra(000189)%>&nbsp;&nbsp;
          <input type="radio" name="vMemb_Active" value="1" <%=fcheck(1, fsqlboolean(vmemb_active))%>><!--webbot bot='PurpleText' PREVIEW='Yes'--><%=fPhra(000024)%>&nbsp; <br>
          <!--webbot bot='PurpleText' PREVIEW='Allows or disallows learner access to this service.'--><%=fPhra(000420)%>&nbsp;
        <!--webbot bot='PurpleText' PREVIEW='To inactive Facilitators or Managers reset learner level to Learner as well.'--><%=fPhra(000421)%>
        </td>
      </tr>
      <% 
        If svMembLevel < 3 Or (svMembLevel = 3 And vCust_MaxUsers > 0) Then 
      %>
      <input type="hidden" name="vMemb_Level" value="<%=vMemb_Level%>">
      <%   
        Else 
      %>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Learner/Security Level'--><%=fPhra(001611)%>:</th>
        <td>

          <%     
          If vCust_ChannelV8 Then 
          %>
          <input type="radio" name="vMemb_Level" value="1" <%=fcheck(1, vmemb_level)%>>1 : <!--webbot bot='PurpleText' PREVIEW='Guest'--><%=fPhra(001643)%> (<!--webbot bot='PurpleText' PREVIEW='can access content and assessments (V8)'--><%=fPhra(001644)%>)<br>
          <% 
            End If 
          %>


          <input type="radio" name="vMemb_Level" value="2" <%=fcheck(2, vmemb_level)%>>2 : <!--webbot bot='PurpleText' PREVIEW='Learner'--><%=fPhra(000165)%> (<!--webbot bot='PurpleText' PREVIEW='can access content and assessments'--><%=fPhra(000093)%>)<br>

          <input type="radio" name="vMemb_Level" value="3" <%=fcheck(3, vmemb_level)%>>3 : <!--webbot bot='PurpleText' PREVIEW='Facilitator'--><%=fPhra(000139)%> (<!--webbot bot='PurpleText' PREVIEW='can add members and monitor progress'--><%=fPhra(000082)%>)<br>
          <%     
          If svMembLevel > 3 Then 
          %>
          <input type="radio" name="vMemb_Level" value="4" <%=fcheck(4, vmemb_level)%>>4 : Manager (can access advanced features)<br>
          <%     
          End If 

          If svMembLevel > 4 Then 
          %>
          <input type="radio" name="vMemb_Level" value="5" <%=fcheck(5, vmemb_level)%>>5 : Administrator 
        <%     
          End If 
        %> 
        </td>
      </tr>
      <% 
        End If 
      %>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Memo'--><%=fPhra(000173)%>: </th>
        <td><input type="text" size="72" name="vMemb_Memo" value="<%=vMemb_Memo%>"></td>
      </tr>

      <% If svMembLevel > 3 Then %>
      <tr>
        <th>V8 Parent | Catalogue: </th>
        <td><%= fV8Fields (svCustAcctId, vMemb_Id)%></td>
      </tr>
      <% 
        End If 
      %>


      <%
        If vCust_MaxSponsor > 0 Then 
            Dim vSponsorList, aSponsors, aSponsors1, aSponsors2
      %>
      <tr>
        <th><% If vMemb_Sponsor > 0 Then %>
          <!--webbot bot='PurpleText' PREVIEW='Sponsor'--><%=fPhra(000489)%>
          <% Else %>
          <!--webbot bot='PurpleText' PREVIEW='Sponsored Learners'--><%=fPhra(000490)%>
          <% End If %> 
        </th>
        <td>
          <% 
            '...get sponsor info
            If vMemb_Sponsor > 0 Then 
              vSponsorList = fSponsorList (vMemb_Sponsor)
              aSponsors = Split(vSponsorList, "|")
          %>
          <a href="User.asp?vMembNo=<%=vMemb_Sponsor%>"><%=aSponsors(0) & " " & aSponsors(1)%></a>
          <% 
            Else
              vSponsorList = fSponsoredList (vMemb_No)
              If Len(vSponsorList) > 0 Then
          %> <a href="Sponsors.asp?vSponsorNo=<%=vMemb_No%>">Edit Sponsored Learners</a><br>&nbsp;
          <table class="table">
          <tr>
            <th align="left"><!--webbot bot='PurpleText' PREVIEW='Name'--><%=fPhra(000187)%></th>
            <th><!--webbot bot='PurpleText' PREVIEW='Expiry Date'--><%=fPhra(000491)%></th>
            <th>&nbsp;</th>
          </tr>
          <%
            aSponsors1 = Split(vSponsorList, "~")
            For i = 0 to Ubound(aSponsors1)
              aSponsors2 = Split(aSponsors1(i), "|")
          %>
          <tr>
            <td><a href="User.asp?vMembNo=<%=aSponsors2(0)%>"><%=aSponsors2(1) & " " & aSponsors2(2)%></a></td>
            <td align="center"><%=aSponsors2(3)%></td>
            <td align="center"><%=fIf(Not aSponsors2(4), "Inactive", "")%></td>
          </tr>
          <%
            Next
          %>
        </table>
          <p>Maximum Sponsored Learners allowed: <%=fIf(vMemb_MaxSponsor=0, vCust_MaxSponsor, vMemb_MaxSponsor)%></p>
          <%  
          Else
          %> 
          (None - <a href="Sponsors.asp?vSponsorNo=<%=vMemb_No%>">Add Sponsored Learners</a>) 
        <%  
          End If
        End If
        %>
        </td>
      </tr>
      <% 
      End If 

  '...display options for criteria - check using fCritOk() 
  '   1) nothing to display since we aren't using criteria
  '   2) select from full criteria for managers+ 
  '   3) select from criteria for facilitators if they manage more than one criteria
  '   4) show criteria to facs if only one in their pervue

  If fCritOk (svCustAcctId) Then

    If svMembLevel > 3 Then  
      i = fCriteriaList (svCustAcctId, "Memb")
      %>
      <tr>
        <th>Group1 Filter :</th>
        <td>
          <select size="<%=vCriteriaListCnt%>" name="vMemb_Criteria" multiple><%=i%></select>
          <br>This is used when we need to assign learners to specific groups (criteria).&nbsp;<font color="#FF0000">Note: Learners must only be assigned to one group</font>&ensp;but Facilitators can be assigned to one or more groups (by using Ctrl+Enter).&nbsp; Managers can be assigned to ALL or multiple groups. </td>
      </tr>
      <%  
    ElseIf Instr(svMembCriteria, " ") > 0 Then 
      i = fCriteriaList (svCustAcctId, "Memb:" & svMembCriteria)      
      %>
      <tr>
        <th>Group1 Filter : </th>
        <td>
          <select size="<%=vCriteriaListCnt%>" name="vMemb_Criteria"><%=i%></select>
          <br>Assign this learner to a specific group. </td>
      </tr>
      <%  
    Else 

      If Not (fNoValue(svMembCriteria) Or svMembCriteria = "0") Then 
      %>
      <tr>
        <th>Group1 Filter : </th>
        <td><%=fCriteria (svMembCriteria)%></td>
      </tr>
      <%   
        End If 
      %>
      <input type="hidden" name="vMemb_Criteria" value="<%=svMembCriteria%>">
      <% 
          End If 

        End If 

        If vCust_Level > 3 Then 
      %>
      <tr>
        <th>Simple Group Filter :</th>
        <td>
          <select size="1" name="vMemb_Group2">
            <% 
            For i = 0 To 24 
              Response.Write "<option " & fSelect(i, vMemb_Group2) & " value='" & i & "'>" & i & "</option>"
            Next 
            %>
          </select>&nbsp;&nbsp; This is a simple &quot;open&quot; filter that can be used for corporate sites.&nbsp; Default is 0 and can range from 1-16.&nbsp; If assigned then the corporate site can use this as a filter to offer certain Programs to different groups, ie Employees (1) vs Managers/Supervisors&nbsp; (2). </td>
      </tr>
      <% 
        Else 
      %>
      <input type="hidden" name="vMemb_Group2" value="<%=vMemb_Group2%>">
      <% 
        End If 
      %>


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
          <select size="6" name="vMemb_Group3">
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





      <%
        '...display options for criteria - check using fCritOk() 
        '   1) nothing to display since we aren't using criteria
        '   2) select from full criteria for managers+ 
        '   3) select from criteria for facilitators if they manage more than one criteria
        '   4) show criteria to facs if only one in their pervue

        '...inactive...
        If fCritOk (svCustAcctId) Then
          If svMembLevel > 3 Then  
            i = fCriteriaList (svCustAcctId, "Memb")
            i = ""
            If Len(i) > 999 Then 
      %>
      <tr>
        <th>Job Title :</th>
        <td><%=i%><br>If blue, then Title was assigned to all jobs within this criteria.&nbsp; If green, then job Title was selected by the learner using the <b>Learning Assessment and Training Plan</b> in <b>My Learning</b> - plus learner may have selected <b>Programs</b> below. </td>
      </tr>
      <% 
            End If
          End If
        End If 



        If svMembLevel > 3 Then 
      %>
      <tr>
        <th>Programs :</th>
        <td>
          <input type="text" size="72" name="vMemb_Programs" value="<%=vMemb_Programs%>" maxlength="8000"><br>
          <b>For channels</b>, enter valid Program Ids separated by spaces, ie: P0011EN P0012EN.&nbsp; These will be added to 'My Content'. &nbsp;&nbsp; NOTE: If adding Programs for channels, you MUST add the expiry date in next field or the duration days in the subsequent field - typically 90 days from today or date of first access.<br><br>
          <b>For corporate</b>, if configured, these values can be generated by the <b>Learning Assessment and Training Plan</b> in <b>My Learning</b>.&nbsp; Note only corporate learners can embrace programs or modules or both in this field, ie: P1234EN 0034EN 0012EN.&nbsp; Values can be entered manually but will be erased whenever learner tries to create his own plan in <b>My Learning</b>.</td>
      </tr>
      <tr>
        <th>Programs Expire :</th>
        <td>
          <input type="text" name="vMemb_Expires" size="20" value="<%=Trim(fFormatSqlDate (vMemb_Expires))%>">
          MMM DD, YYYY (ie: Jan 1, 2004)<br>If entered, signifies date that the above programs (and/or modules) expire, ie <% =fFormatSqlDate(Now + 90)%>.</td>
      </tr>
      <tr>
        <th>Programs Duration :</th>
        <td>
          <input type="text" name="vMemb_Duration" size="6" value="<%=vMemb_Duration%>" maxlength="3">
          Specify in days. Typically used when the above expiry field is dependent on when learner first visits the site - computed on signin as current date plus duration days.&nbsp; This is only used if the above field is initially found empty on signin.&nbsp; Once the expiry date is computed, this field is no longer used (except as a memo).&nbsp; This is useful issuing Access Ids on consignment.</td>
      </tr>
      <tr>
        <th>First Visit : </th>
        <td>
          <input type="text" name="vMemb_FirstVisit" size="20" value="<%=fFormatSqlDate (vMemb_FirstVisit)%>">
          ie <% =fFormatSqlDate(Now)%>.<br>Do not leave empty or it will revert to today&#39;s date. </td>
      </tr>

<!--      <tr>
        <th>Authoring Rights :</th>
        <td>
          <input type="radio" name="vMemb_Auth" value="1" <%=fcheck(fsqlboolean(vmemb_auth), 1)%>>YES&nbsp;&nbsp;&nbsp;
          <input type="radio" name="vMemb_Auth" value="0" <%=fcheck(fsqlboolean(vmemb_auth), 0)%>>NO<br>Select YES if this facilitator or manager is authorized to use VuBuild. Only valid if the Customer is authorized to use VuBuild. </td>
      </tr>-->



      <% 
        Else 
      %>
      <input type="hidden" name="vMemb_Programs" value="<%=vMemb_Programs%>">
      <input type="hidden" name="vMemb_Expires" value="<%=vMemb_Expires%>"><input type="hidden" name="vMemb_Duration" value="<%=vMemb_Duration%>"><input type="hidden" name="vMemb_FirstVisit" value="<%=vMemb_FirstVisit%>">
      <input type="hidden" name="vMemb_Auth" value="<%=fSqlBoolean(vMemb_Auth)%>">
      <% 
        End If 


          '...Display Jobs unless they are all mandatory from the criteria table
          i = fIf (fDisplayJobs, Trim(fJobsProgsOptions (vMemb_Jobs)), "")
          If svMembLevel > 2 And Len(i) > 0 Then
      %>
      <tr>
        <th nowrap>Training Path :</th>
        <td>
          <select size="<%=vJobsListCnt%>" name="vMemb_Jobs" multiple style="font-family: Lucida Console"><%=i%></select>
          <br>Select one or more of Training Paths (Job Streams) for <b>My Learning</b>.<br>Use CTRL+Enter to select more than one Job Paths. </td>
      </tr>
      <%
        Else 
      %>
      <input type="hidden" name="vMemb_Jobs" value="<%=vMemb_Jobs%>">
      <% 
        End If 

        If svMembLevel > 4 Then 
      %>
      <tr>
        <th height="56">VuNews :</th>
        <td height="56">
          <input type="radio" name="vMemb_VuNews" value="1" <%=fcheck(fsqlboolean(vmemb_vunews), 1)%>>YES&nbsp;&nbsp;&nbsp;
          <input type="radio" name="vMemb_VuNews" value="0" <%=fcheck(fsqlboolean(vmemb_vunews), 0)%>>NO<br>Select YES if this learner is to receive VuNews (English only).&nbsp; Defaults to YES if the configured as ON on the Customer Table.&nbsp; This can also be turned off by the learner on the Info page if that page is enabled. </td>
      </tr>
      <% 
  Else 
      %>
      <input type="hidden" name="vMemb_VuNews" value="<%=fSqlBoolean(vMemb_VuNews)%>"><% 
  End If 

 
  If svMembLevel > 4 Then 
      %>
      <tr>
        <th colspan="2" style="text-align:center;"><br><br>Special Manager Rights - Ensure you only assign these to Managers!<br><br></th>
      </tr>
      <tr class="notice">
        <th>My Learning :</th>
        <td>
          <input type="radio" name="vMemb_MyWorld" value="1" <%=fcheck(fsqlboolean(vmemb_myworld), 1)%>>YES&nbsp;&nbsp;&nbsp;
          <input type="radio" name="vMemb_MyWorld" value="0" <%=fcheck(fsqlboolean(vmemb_myworld), 0)%>>NO<br>Select YES if this manager can edit My Learning. </td>
      </tr>
      <tr class="notice">
        <th>LCMS :</th>
        <td>
          <input type="radio" name="vMemb_LCMS" value="1" <%=fcheck(fsqlboolean(vmemb_lcms), 1)%>>YES&nbsp;&nbsp;&nbsp;
          <input type="radio" name="vMemb_LCMS" value="0" <%=fcheck(fsqlboolean(vmemb_lcms), 0)%>>NO<br>Select YES if this manager can edit content tables: ie programs, modules, tests, exams. </td>
      </tr>

<!--      <tr class="notice">
        <th>VuBuild :</th>
        <td>
          <input type="radio" name="vMemb_VuBuild" value="1" <%=fcheck(fsqlboolean(vmemb_vubuild), 1)%>>YES&nbsp;&nbsp;&nbsp;
          <input type="radio" name="vMemb_VuBuild" value="0" <%=fcheck(fsqlboolean(vmemb_vubuild), 0)%>>NO<br>Select YES if this manager can create and delete VuBuild customers</td>
      </tr>
-->


      <tr class="notice">
        <th>Manual Ecom :</th>
        <td>
          <input type="radio" name="vMemb_Ecom" value="1" <%=fcheck(fsqlboolean(vmemb_ecom), 1)%>>YES&nbsp;&nbsp;&nbsp;
          <input type="radio" name="vMemb_Ecom" value="0" <%=fcheck(fsqlboolean(vmemb_ecom), 0)%>>NO<br>Select YES if this manager can post ecommerce transactions that bypass Internet Secure. </td>
      </tr>
      <tr class="notice">
        <th height="46">Super Manager :</th>
        <td height="46">
          <input type="radio" name="vMemb_Manager" value="1" <%=fcheck(fsqlboolean(vmemb_manager), 1)%>>YES&nbsp;&nbsp;&nbsp;
          <input type="radio" name="vMemb_Manager" value="0" <%=fcheck(fsqlboolean(vmemb_manager), 0)%>>NO<br>Select YES if this manager can have advanced ecommerce rights, ie access expired accounts, etc.</td>
      </tr>
      <tr>
        <th colspan="2" class="overline" align="center"></th>
      </tr>
      <% 
        Else 
      %>
      <input type="hidden" name="vMemb_MyWorld" value="<%=fSqlBoolean(vMemb_MyWorld)%>">
      <input type="hidden" name="vMemb_LCMS" value="<%=fSqlBoolean(vMemb_LCMS)%>">
      <input type="hidden" name="vMemb_Channel" value="<%=fSqlBoolean(vMemb_Channel)%>">
      <input type="hidden" name="vMemb_VuBuild" value="<%=fSqlBoolean(vMemb_VuBuild)%>">
      <input type="hidden" name="vMemb_Ecom" value="<%=fSqlBoolean(vMemb_Ecom)%>">
      <input type="hidden" name="vMemb_Manager" value="<%=fSqlBoolean(vMemb_Manager)%>">
      <% 
        End If 
      %>


      <tr>
        <td style="text-align:center" colspan="2"><br>
          <% 
          If Len(vNext) > 0 Then 
          %> <%=f5%><input onclick="location.href = '<%=fjUnQuote(vNext)%>'" type="button" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button085"><%=f5%> <% 
          End If 

        If svCustAcctId <> fDefault(vMemb_AcctId, svCustAcctId) Then 
        %>
          <h5><br><br>
            <!--webbot bot='PurpleText' PREVIEW='Learner Profiles accessed from another Account are Read Only.'--><%=fPhra(001265)%>
            <br>
            <!--webbot bot='PurpleText' PREVIEW='You cannot Update or Delete this Learner&#39;s Profile'--><%=fPhra(001572)%>.<br></h5>
          <% 
            Else
  
              If ((svMembLevel = 4 And vCust_DeleteLearners And vCust_MaxUsers = 0) Or svMembManager Or svMembLevel = 5) Then 
          %> <%=f10%>
          <input onclick="jconfirm('User<%=fGroup%>.asp?vDelete=<%=vMemb_No%>&amp;vNext=<%=Server.UrlEncode(vNext)%>', '<!--webbot bot='PurpleText' PREVIEW='Ok to delete?'--><%=fPhra(000199)%>')" type="button" value="<%=bDelete%>" name="bDelete" class="button085">
          <%=f5%> 
          <%   
              End If 

              If ((svMembLevel < 5 And vCust_InsertLearners And vCust_UpdateLearners) Or svMembManager Or svMembLevel = 5) Then 
                If svMembLevel > 3 Or vCust_MaxUsers = 0 Or (vCust_MaxUsers > 0 And (vCust_MaxUsers - fAllMembCount) > -3) Then
                    %> <%=f5%><input type="submit" value="<%=bUpdate%>" name="bUpdate" class="button085">
                    <%     
                End If
              End If 

            End If 
          %> 
          <br>
          <h2><a href="Users_o.asp?vSort=id&vStart=<%=vMemb_Id%>&vLastValue=<%=vMemb_LastName & vMemb_FirstName & vMemb_No%>"><!--webbot bot='PurpleText' PREVIEW='Learner Report'--><%=fPhra(000367)%></a></h2>
          <h2 align="center"><%=vCust_Id & "  (" & vCust_Title & ")"%></h2>
        </td>
      </tr>
    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


