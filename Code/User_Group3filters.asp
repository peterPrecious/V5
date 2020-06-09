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

    If fNoValue(vMemb_Id) Then
      vMessage = fPhraH(001212)
    Else
     vMembId = Request("vMembId")
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
  <meta charset="UTF-8">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>User Profiles</title>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <script>
    function Validate(theForm)
    {
   
      if (theForm.vMemb_Id.value == "")
      {
        alert("Please enter a value for the \"Id / Password\" field.");
        theForm.vMemb_Id.focus();
        return (false);
      }
    
      if (theForm.vMemb_Id.value.length < 4)
      {
        alert("Please enter at least 4 characters in the \"Id / Password\" field.");
        theForm.vMemb_Id.focus();
        return (false);
      }
    
      if (theForm.vMemb_Id.value.length > 64)
      {
        alert("Please enter at most 64 characters in the \"Id / Password\" field.");
        theForm.vMemb_Id.focus();
        return (false);
      }
    
      var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_-@.";
      var checkStr = theForm.vMemb_Id.value;
      var allValid = true;
      for (i = 0;  i < checkStr.length;  i++)
      {
        ch = checkStr.charAt(i);
        for (j = 0;  j < checkOK.length;  j++)
          if (ch == checkOK.charAt(j))
            break;
        if (j == checkOK.length)
        {
          allValid = false;
          break;
        }
      }
      if (!allValid)
      {
        alert("Please enter only letters, digits and \"_-@.\" characters in the \"Id / Password\" field.");
        theForm.vMemb_Id.focus();
        return (false);
      }


      //  only check password if used by this memeber, else ignore
      if (theForm.vPassword.value == "check" && theForm.vMemb_Level.value == 2) 
      {
      
        if (theForm.vMemb_Pwd.value == "")
        {
          alert("Please enter a value for the \"Password\" field.");
          theForm.vMemb_Pwd.focus();
          return (false);
        }
      
        if (theForm.vMemb_Pwd.value.length < 4)
        {
          alert("Please enter at least 4 characters in the \"Password\" field.");
          theForm.vMemb_Pwd.focus();
          return (false);
        }
      
        if (theForm.vMemb_Pwd.value.length > 64)
        {
          alert("Please enter at most 64 characters in the \"Password\" field.");
          theForm.vMemb_Pwd.focus();
          return (false);
        }
      
        var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789-_-@.";
        var checkStr = theForm.vMemb_Pwd.value;
        var allValid = true;
        for (i = 0;  i < checkStr.length;  i++)
        {
          ch = checkStr.charAt(i);
          for (j = 0;  j < checkOK.length;  j++)
            if (ch == checkOK.charAt(j))
              break;
          if (j == checkOK.length)
          {
            allValid = false;
            break;
          }
        }
        if (!allValid)
        {
          alert("Please enter only letter, digit and \"_-@.\" characters in the \"Password\" field.");
          theForm.vMemb_Pwd.focus();
          return (false);
        }
  
      }

    
      if (theForm.vMemb_Duration.value.length > 3)
      {
        alert("Please enter at most 3 characters in the \"Duration (1-365)\" field.");
        theForm.vMemb_Duration.focus();
        return (false);
      }
    
      var checkOK = "0123456789";
      var checkStr = theForm.vMemb_Duration.value;
      var allValid = true;
      var decPoints = 0;
      var allNum = "";
      for (i = 0;  i < checkStr.length;  i++)
      {
        ch = checkStr.charAt(i);
        for (j = 0;  j < checkOK.length;  j++)
          if (ch == checkOK.charAt(j))
            break;
        if (j == checkOK.length)
        {
          allValid = false;
          break;
        }
        allNum += ch;
      }
      if (!allValid)
      {
        alert("Please enter only digit characters in the \"Duration (1-365)\" field.");
        theForm.vMemb_Duration.focus();
        return (false);
      }


     if (theForm.vMemb_Criteria != undefined) {

       if (theForm.vMemb_Criteria.selectedIndex < 0)
       {
         alert("Please select one of the \"Group\" options.");
         theForm.vMemb_Criteria.focus();
         return (false);
       }            
     }

      return (true);
    }
  </script>  
  
      
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>

  <form name="FrontPage_Form1" method="POST" action="User<%=fGroup%>.asp" target="_self" onsubmit="return Validate(this)">

    <input type="hidden" name="vHidden"  value="Y">
    <input type="hidden" name="vNext"    value="<%=vNext%>">
    <input type="hidden" name="vMemb_No" value="<%=vMemb_No%>">
    <input type="hidden" name="vMembId"  value="<%=vMemb_Id%>">  <!-- this is the ID before update - incase we try to change it to an existing ID -->

    <table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td width="100%" valign="top" colspan="2" align="center">
        <h1>
        <!--webbot bot='PurpleText' PREVIEW='Add a Learner'--><%=fPhra(000370)%> /&nbsp;
        <!--webbot bot='PurpleText' PREVIEW='Learner Profile'--><%=fPhra(000371)%></h1>

        <% 
          If vCust_MaxUsers > 0 And (vCust_Level < 3 Or svMembLevel = 5 Or svMembManager) Then 
            p0 = vCust_MaxUsers
            p1 = fAllMembCount - 1
        %>

        <h6 align="left">
        <!--webbot bot='PurpleText' PREVIEW='Note: This account is limited to ^0 active or inactive learners.&nbsp; Once you reach the maximum you will be unable to add new learners or edit existing learners.&nbsp; You currently have ^1 learners on file.'--><%=fPhra(000512)%></h6>
        <% End If %> 
        
        <% If Not fNoValue(vMessage) Then %>
        <h5><%=vMessage%></h5>
        <% End If %> 
        
        </td>
      </tr>

      <tr>

        <th align="right" width="25%" valign="top"><%=fIf(svCustPwd, fPhraH(000374), fPhraH(000211))%> :</th>

        <td width="75%" valign="top">

          <% If svMembLevel < 3 Then %>     
            <%=vMemb_Id%>
            <% = fIf (svMembLevel = 5, f10 & "(Learner No : " & vMemb_No & ")", "") %>
            <input type="hidden" name="vMemb_Id" value="<%=vMemb_Id%>">

          <% ElseIf InStr(vMemb_Id, vPasswordx) > 0 Then %>
            **********                      
            <% = fIf (svMembLevel = 5, f10 & "(Learner No : " & vMemb_No & ")", "") %>
            <input type="hidden" name="vMemb_Id" value="<%=vMemb_Id%>">

          <% Else %>
            <input type="text" size="30" name="vMemb_Id" value="<%=vMemb_Id%>" maxlength="64" class="c2">
            <% = fIf (svMembLevel = 5, f10 & "(Learner No : " & vMemb_No & ")", "") %>
            <br><!--webbot bot='PurpleText' PREVIEW='Must be unique using only English alpha, numeric and &quot;_.-@&quot; characters.'--><%=fPhra(000372)%>

          <% End If %>

        </td>

      </tr>



      
      <% If vCust_Pwd Then %>
      <tr>
        <th align="right" width="25%" valign="top">
        <!--webbot bot='PurpleText' PREVIEW='Password'--><%=fPhra(000211)%> :</th>
        <td width="75%" valign="top">
          <input type="text" size="30" name="vMemb_Pwd" value="<%=vMemb_Pwd%>" maxlength="64" class="c2"><br><!--webbot bot='PurpleText' PREVIEW='Assigned by learner using only English alpha, numeric and &quot;_.-@&quot; characters.'--><%=fPhra(000419)%>
          <br><font color="#FF0000">Do not use for facilitators or managers.</font><
        /td>
      </tr>
      <input type="hidden" name="vPassword" value="check">
      <% Else %>
      <input type="hidden" name="vPassword" value="ignore">
      <% End If %>


      <tr>
        <th align="right" width="25%" valign="top">
        <!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%> :</th>
        <td width="75%" valign="top"><input type="text" size="30" name="vMemb_FirstName" value="<%=vMemb_FirstName%>" maxlength="32" class="c2"></td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top">
        <!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%> :</th>
        <td width="75%" valign="top"><input type="text" size="30" name="vMemb_LastName" value="<%=vMemb_LastName%>" maxlength="64" class="c2"></td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top">
        <!--webbot bot='PurpleText' PREVIEW='Email Address'--><%=fPhra(000126)%> :</th>
        <td width="75%" valign="top"><input type="text" size="46" name="vMemb_Email" value="<%=vMemb_Email%>" maxlength="128" class="c2"></td>
      </tr>

      <tr>
        <th align="right" width="25%" valign="top">
        <!--webbot bot='PurpleText' PREVIEW='Organization'--><%=fPhra(000470)%> :</th>
        <td width="75%" valign="top"><input type="text" size="46" name="vMemb_Organization" value="<%=vMemb_Organization%>" maxlength="128" class="c2"></td>
      </tr>


      <tr>
        <th align="right" width="25%" valign="top">
          <!--webbot bot='PurpleText' PREVIEW='Active'--><%=fPhra(000063)%> :</th>
        <td width="75%" valign="top">
          <% If fNoValue(vMemb_Active) Then vMemb_Active = 1 %>
          <input type="radio" name="vMemb_Active" value="0" <%=fcheck(0, fsqlboolean(vmemb_active))%>><!--webbot bot='PurpleText' PREVIEW='No'--><%=fPhra(000189)%>&nbsp;&nbsp; 
          <input type="radio" name="vMemb_Active" value="1" <%=fcheck(1, fsqlboolean(vmemb_active))%>><!--webbot bot='PurpleText' PREVIEW='Yes'--><%=fPhra(000024)%>&nbsp; <br>
          <!--webbot bot='PurpleText' PREVIEW='Allows or disallows learner access to this service.'--><%=fPhra(000420)%>&nbsp;
          <!--webbot bot='PurpleText' PREVIEW='To inactive Facilitators or Managers reset learner level to Learner as well.'--><%=fPhra(000421)%>
        </td>
      </tr>


      <% If svMembLevel < 3 Or (svMembLevel = 3 And vCust_MaxUsers > 0) Then %> 
      <input type="hidden" name="vMemb_Level" value="<%=vMemb_Level%>">
      <% Else %>
      <tr>
        <th align="right" width="25%" valign="top">
        <!--webbot bot='PurpleText' PREVIEW='Learner Level'--><%=fPhra(000373)%> :</th>
        <td width="75%" valign="top">
          <input type="radio" name="vMemb_Level" value="2" <%=fcheck(2, vmemb_level)%>>2: <!--webbot bot='PurpleText' PREVIEW='Learner'--><%=fPhra(000165)%> (<!--webbot bot='PurpleText' PREVIEW='can access content and assessments'--><%=fPhra(000093)%>)<br>
          <input type="radio" name="vMemb_Level" value="3" <%=fcheck(3, vmemb_level)%>>3: <!--webbot bot='PurpleText' PREVIEW='Facilitator'--><%=fPhra(000139)%> (<!--webbot bot='PurpleText' PREVIEW='can add members and monitor progress'--><%=fPhra(000082)%>)<br>
          <% If svMembLevel > 3 Then %>
          <input type="radio" name="vMemb_Level" value="4" <%=fcheck(4, vmemb_level)%>>4: Manager (can access advanced features)<br>
          <% End If %>
          <% If svMembLevel > 4 Then %>
          <input type="radio" name="vMemb_Level" value="5" <%=fcheck(5, vmemb_level)%>>5: Administrator 
          <% End If %> 
        </td>
      </tr>
      <% End If %> 
      
      <tr>
        <th align="right" width="25%">
        <!--webbot bot='PurpleText' PREVIEW='Memo'--><%=fPhra(000173)%> :</th>
        <td width="75%"><input type="text" size="72" name="vMemb_Memo" value="<%=vMemb_Memo%>" class="c2"></td>
      </tr>

      <%
        If vCust_MaxSponsor > 0 And vMemb_No > 0 Then 
          Dim vSponsorList, aSponsors, aSponsors1, aSponsors2
      %>
      
      <tr>

        <th align="right" width="25%" valign="top">
          <% = fIf(vMemb_Sponsor > 0, "<!--webbot bot='PurpleText' PREVIEW='Sponsor'--><%=fPhra(000489)%>", "<!--webbot bot='PurpleText' PREVIEW='Sponsored Learners'--><%=fPhra(000490)%>") %>
        </th>

        <td width="75%">
          <% '...get sponsor info
             If vMemb_Sponsor > 0 Then 
               vSponsorList = fSponsorList (vMemb_Sponsor)
               aSponsors = Split(vSponsorList, "|")
          %>
             <a href="User.asp?vMembNo=<%=vMemb_Sponsor%>"><%=aSponsors(0) & " " & aSponsors(1)%></a>
          <% 
             Else
               vSponsorList = fSponsoredList (vMemb_No)
               If Len(vSponsorList) > 0 Then
          %>
                 <a href="Sponsors.asp?vSponsorNo=<%=vMemb_No%>">Edit Sponsored Learners</a><br>&nbsp;<table border="0" id="table1" cellspacing="0" cellpadding="2">
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
                 (None -
                 <a href="Sponsors.asp?vSponsorNo=<%=vMemb_No%>">Add Sponsored Learners</a>) 
        <%  
               End If
             End If
        %>
        </td>

      </tr>
      <% End If %>



      <% 

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
        <th align="right" width="25%" valign="top">Group1 Filter :
        <!--
        <br>
        [coming: for assigning content] 
        -->
        </th>
        <td width="75%">
          <select size="<%=vCriteriaListCnt%>" name="vMemb_Criteria" multiple class="c2">
          <%=i%>
          </select>
          <br>This is used when we need to assign learners to specific groups (criteria).&nbsp; <font color="#FF0000">Note: Learners must only be assigned to one group</font> but Facilitators can be assigned to one or more groups (by using Ctrl+Enter).&nbsp; Managers can be assigned to ALL or multiple groups.</td>
      </tr>

      <%  
          ElseIf Instr(svMembCriteria, " ") > 0 Then 
            i = fCriteriaList (svCustAcctId, "Memb:" & svMembCriteria)      
      %> 

      <tr>
        <th align="right" width="25%" valign="top">Group1 Filter :
        <!--
         <br>
        [coming: for assigning content] 
        -->
        </th>
        <td width="75%">
          <select size="<%=vCriteriaListCnt%>" name="vMemb_Criteria"><%=i%></select>
          <br>Assign this learner to a specific group.
        </td>
      </tr>

      <%  Else %> 

      <%   If Not (fNoValue(svMembCriteria) Or svMembCriteria = "0") Then %>
      <tr>
        <th align="right" width="25%">Group1 Filter :
        <!--
         <br>
        [coming: for assigning content] 
        -->
        </th>
        <td width="75%"><%=fCriteria (svMembCriteria)%></td>
      </tr>
      <%   End If %> 

      <input type="hidden" name="vMemb_Criteria" value="<%=svMembCriteria%>">

      <% 
          End If 

        End If 
      %> 
      

      <% If vCust_Level > 2 Then %>
      <tr>
        <th align="right" width="25%" valign="top">Group2 Filter :</th>
        <td width="75%">
        <select size="1" name="vMemb_Group2" class="c2">
          <% 
            For i = 0 To 24 
              Response.Write "<option " & fSelect(i, vMemb_Group2) & " value='" & i & "'>" & i & "</option>"
            Next 
          %>
        </select>&nbsp; This is a simple &quot;open&quot; filter that can be used for corporate sites.&nbsp; Default is 0 and can range from 1-16.&nbsp; If assigned then the corporate site can use this as a filter to offer certain Programs to different groups, ie Employees (1) vs Managers/Supervisors&nbsp; (2).</td>
      </tr>
      <% Else %> 
      <input type="hidden" name="vMemb_Group2" value="<%=vMemb_Group2%>">
      <% End If %>













      <% 

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
        <th align="right" width="25%" valign="top">Group3 Filter :<br>
        [coming: for assigning reporting] </th>
        <td width="75%">
          <select size="<%=vCriteriaListCnt%>" name="vMemb_Report" multiple class="c2">
          <%=i%>
          </select>
          [DO NOT USE - COMING]<br>This is like the Group 1 filter above but it use to assign Groups to Facilitators for Reporting.&nbsp; It does NOT affect any assignment of learning.&nbsp; Managers can be assigned to ALL or multiple groups.</td>
      </tr>

      <%  
          ElseIf Instr(svMembCriteria, " ") > 0 Then 
            i = fCriteriaList (svCustAcctId, "Memb:" & svMembCriteria)      
      %> 

      <tr>
        <th align="right" width="25%" valign="top">Group3 Filter :<br>
        [coming: for assigning reporting] </th>
        <td width="75%">
          <select size="<%=vCriteriaListCnt%>" name="vMemb_Report"><%=i%></select>
          <br>Assign this learner to a specific group.
        </td>
      </tr>

      <%  Else %> 

      <%   If Not (fNoValue(svMembCriteria) Or svMembCriteria = "0") Then %>
      <tr>
        <th align="right" width="25%">Group3 Filter :<br>
        [coming: for assigning reporting] </th>
        <td width="75%"><%=fCriteria (svMembCriteria)%></td>
      </tr>
      <%   End If %> 




      <input type="hidden" name="vMemb_Criteria" value="<%=svMembCriteria%>">

      <% 
          End If 

        End If 
      %> 
      





















      <% '...display job title if available ===  inactive
       i = ""
'      i = fJobsTitleByNo(vMemb_JobsNo) 
'      If Len(i) > 0 Then 

       If Len(i) > 999 Then 
      %>
      <tr>
        <th align="right" width="25%" valign="top">Job Title :</th>
        <td width="75%"><%=i%><br>If blue, then Title was assigned to all jobs within this criteria.&nbsp; If green, then job Title was selected by the learner using the <b>Learning Assessment and Training Plan</b> in <b>My Learning</b> - plus learner may have selected <b>Programs</b> below.</td>
      </tr>
      <% End If %> 

     
      <% If svMembLevel > 3 Then %>
      <tr>
        <th align="right" width="25%" valign="top">
        Programs :</th>
        <td width="75%" valign="top"><input type="text" size="72" name="vMemb_Programs" value="<%=vMemb_Programs%>" maxlength="8000" class="c2"><br><b>For channels</b>, enter valid programs separated by spaces, ie: P0011EN P0012EN.&nbsp; This add these programs to If you enter a <b>Start at</b> value, a partial report will display learners from that value to the report&#39;s end.&nbsp; Normal program source is the Customer program string plus any ecommerce purchases.&nbsp;&nbsp; NOTE: If adding Programs for channels, you MUST add the expiry date in next field or the duration days in the subsequent field - typically 90 days from today or date of first access.<br><br><b>For corporate</b>, if configured, these values can be generated by the <b>Learning Assessment and Training Plan</b> in <b>My Learning</b>.&nbsp; Note only corporate learners can embrace programs or modules or both in this field, ie: P1234EN 0034EN 0012EN.&nbsp; Values can be entered manually but will be erased whenever learner tries to create his own plan in <b>My Learning</b>.</td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top">
        Programs Expire :</th>
        <td width="75%" valign="top"><input type="text" name="vMemb_Expires" size="20" value="<%=Trim(fFormatSqlDate (vMemb_Expires))%>"> MMM DD, YYYY (ie: Jan 1, 2004)<br>If entered, signifies date that the above programs (and/or modules) expire, ie <% =fFormatSqlDate(Now + 90)%>.</td>
      </tr>
      <tr>
        <th align="right" width="25%" valign="top">Programs Duration :</th>
        <td width="75%" valign="top"><input type="text" name="vMemb_Duration" size="6" value="<%=vMemb_Duration%>" maxlength="3"> Specify in days. Typically used when the above expiry field is dependent on when learner first visits the site - computed on signin as current date plus duration days.&nbsp; This is only used if the above field is initially found empty on signin.&nbsp; Once the expiry date is computed, this field is no longer used (except as a memo).&nbsp; This is useful issuing Access Ids on consignment.</td>
      </tr>

      <% Else %> 
      <input type="hidden" name="vMemb_Programs" value="<%=vMemb_Programs%>">
      <input type="hidden" name="vMemb_Expires"  value="<%=vMemb_Expires%>">
      <input type="hidden" name="vMemb_Duration" value="<%=vMemb_Duration%>">
      <% End If %> 

      <% 
       '...Display Jobs unless they are all mandatory from the criteria table
       If fDisplayJobs Then 
         i = Trim(fJobsProgsOptions (vMemb_Jobs)) 
       Else
         i = ""
       End If

       If svMembLevel > 2 And Len(i) > 0 Then
      %>
      <tr>
        <th valign="top" nowrap align="right" width="25%">Training Path :</th>
        <td width="75%">
          <select size="<%=vJobsListCnt%>" name="vMemb_Jobs" multiple style="font-family: Lucida Console" class="c2">
          <%=i%>
          </select> 
          <br>Select one or more of Training Paths (Job Streams) for <b>My Learning</b>.<br>Use CTRL+Enter to select more than one Job Paths. 
        </td>
      </tr>
      <%
         Else %> 
      <input type="hidden" name="vMemb_Jobs" value="<%=vMemb_Jobs%>">
      <% End If %>



      <% If svMembLevel > 3 Then %>
      <tr>
        <th align="right" width="25%" valign="top">First Visit : </th>
        <td width="75%" valign="top"><input type="text" name="vMemb_FirstVisit" size="20" value="<%=fFormatSqlDate (vMemb_FirstVisit)%>" class="c2"> ie <% =fFormatSqlDate(Now)%>.<br>
        Do not leave empty or it will revert to today's date.</td>
      </tr>
      <% Else %> 
      <input type="hidden" name="vMemb_FirstVisit" value="<%=vMemb_FirstVisit%>">
      <% End If %> 

      <% If svMembLevel > 3 Then %>
      <tr>
        <th align="right" valign="Top" width="25%" height="45">Authoring Rights :</th>
        <td valign="Top" width="75%" height="45"><input type="radio" name="vMemb_Auth" value="1" <%=fcheck(fsqlboolean(vmemb_auth), 1)%>>YES&nbsp;&nbsp;&nbsp; <input type="radio" name="vMemb_Auth" value="0" <%=fcheck(fsqlboolean(vmemb_auth), 0)%>>NO<br>Select YES if this manager is authorized to use VuBuild. Only valid if the Customer is authorized to use VuBuild.</td>
      </tr>
      <% Else %> 
      <input type="hidden" name="vMemb_Auth" value="<%=fSqlBoolean(vMemb_Auth)%>">
      <% End If %> 

      <% If svMembLevel > 4 Then %>
      <tr>
        <th align="right" valign="Top" width="25%" height="56">VuNews :</th>
        <td valign="Top" width="75%" height="56">
          <input type="radio" name="vMemb_VuNews" value="1" <%=fcheck(fsqlboolean(vmemb_VuNews), 1)%>>YES&nbsp;&nbsp;&nbsp; 
          <input type="radio" name="vMemb_VuNews" value="0" <%=fcheck(fsqlboolean(vmemb_VuNews), 0)%>>NO<br>Select YES if this learner is to receive VuNews (English only).&nbsp; Defaults to YES if the configured as ON on the Customer Table.&nbsp; This can also be turned off by the learner on the Info page if that page is enabled.</td>
      </tr>
      <% Else %> 
      <input type="hidden" name="vMemb_VuNews" value="<%=fSqlBoolean(vMemb_VuNews)%>">
      <% End If %> 

      <% If spGapPosnsExistByAcctId (svCustAcctId) Then %>

      <%   If svMembLevel > 4 Then %>
      <tr>
        <th align="right" valign="Top" width="25%">GAP Position No :</th>
        <td valign="Top" width="75%">
          <input type="text" name="vMemb_GapPositionNo" size="4" value="<%=vMemb_GapPositionNo%>" class="c2"><%="&nbsp;&nbsp;&nbsp;" & spPosnTitleByNo (vMemb_GapPositionNo) %>
          <br>This value is created when the GAP Manager uploads the Organization's Position Codes. It points to this learner's Position Code in the <a class="c2" target="_blank" href="Gap_TablePositions.asp">Position Table</a>.&nbsp; A learner cannot participate in the GAP service without a valid (non zero) Position No.
        </td>
      </tr>
      <%   Else %>
      <tr>
        <th align="right" valign="Top" width="25%">GAP Position No :</th>
        <td valign="Top" width="75%">
          <b><%=vMemb_GapPositionNo & "&nbsp;&nbsp;&nbsp;" & spPosnTitleByNo (vMemb_GapPositionNo) %></b>
          <br>This value is created when the GAP Manager uploads the Organization's Position Codes. It points to this learner's Position Code in the Posn Table.&nbsp; A learner cannot participate in the GAP service without a valid (non zero) Position No.
        </td>
      </tr>
      <input type="hidden" name="vMemb_GapPositionNo" value="<%=vMemb_GapPositionNo%>">
      <%   End If %>

      <% End If %>
   

      <% If svMembLevel > 4 Then %>
      <tr>
        <th valign="Top" width="100%" colspan="2"><hr noshade color="#DDEEF9" size="5"> 
        <br>Special Manager Rights<br>&nbsp;</th>
      </tr>
      <tr>
        <th align="right" valign="Top" width="25%">My Learning :</th>
        <td valign="Top" width="75%"><input type="radio" name="vMemb_MyWorld" value="1" <%=fcheck(fsqlboolean(vmemb_myworld), 1)%>>YES&nbsp;&nbsp;&nbsp; <input type="radio" name="vMemb_MyWorld" value="0" <%=fcheck(fsqlboolean(vmemb_myworld), 0)%>>NO<br>Select YES if this manager can edit My Learning. </td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="25%">LCMS :</th>
        <td valign="Top" width="75%"><input type="radio" name="vMemb_LCMS" value="1" <%=fcheck(fsqlboolean(vmemb_lcms), 1)%>>YES&nbsp;&nbsp;&nbsp; <input type="radio" name="vMemb_LCMS" value="0" <%=fcheck(fsqlboolean(vmemb_lcms), 0)%>>NO<br>Select YES if this manager can edit content tables: ie programs, modules, tests, exams. </td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="25%">VuBuild :</th>
        <td valign="Top" width="75%"><input type="radio" name="vMemb_VuBuild" value="1" <%=fcheck(fsqlboolean(vmemb_vubuild), 1)%>>YES&nbsp;&nbsp;&nbsp; <input type="radio" name="vMemb_VuBuild" value="0" <%=fcheck(fsqlboolean(vmemb_vubuild), 0)%>>NO<br>Select YES if this manager can create and delete VuBuild customers (Accounts 7xxx). </td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="25%">Manual Ecom :</th>
        <td valign="Top" width="75%"><input type="radio" name="vMemb_Ecom" value="1" <%=fcheck(fsqlboolean(vmemb_ecom), 1)%>>YES&nbsp;&nbsp;&nbsp; <input type="radio" name="vMemb_Ecom" value="0" <%=fcheck(fsqlboolean(vmemb_ecom), 0)%>>NO<br>Select YES if this manager can post ecommerce transactions that bypass Internet Secure. </td>
      </tr>
      <tr>
        <th align="right" valign="Top" width="25%" height="46">Super Manager :</th>
        <td valign="Top" width="75%" height="46"><input type="radio" name="vMemb_Manager" value="1" <%=fcheck(fsqlboolean(vmemb_manager), 1)%>>YES&nbsp;&nbsp;&nbsp; <input type="radio" name="vMemb_Manager" value="0" <%=fcheck(fsqlboolean(vmemb_manager), 0)%>>NO<br>Select YES if this manager can have advanced ecommerce rights, ie access expired accounts, etc.</td>
      </tr>

      <% Else %> 

      <input type="hidden" name="vMemb_MyWorld" value="<%=fSqlBoolean(vMemb_MyWorld)%>">
      <input type="hidden" name="vMemb_LCMS"    value="<%=fSqlBoolean(vMemb_LCMS)%>">
      <input type="hidden" name="vMemb_Channel" value="<%=fSqlBoolean(vMemb_Channel)%>">
      <input type="hidden" name="vMemb_VuBuild" value="<%=fSqlBoolean(vMemb_VuBuild)%>">
      <input type="hidden" name="vMemb_Ecom"    value="<%=fSqlBoolean(vMemb_Ecom)%>">
      <input type="hidden" name="vMemb_Manager" value="<%=fSqlBoolean(vMemb_Manager)%>">

      <% End If %> 


      <% If vCust_Tab4 And vCust_Tab4Type = "GA" And svMembLevel > 3 And vMemb_Level > 2 Then %>
      <tr>
        <th align="right" valign="Top" width="25%" height="46">GAP :</th>
        <td valign="Top" width="75%" height="46"><input type="radio" name="vMemb_GAP" value="1" <%=fcheck(fsqlboolean(vmemb_gap), 1)%>>YES&nbsp;&nbsp;&nbsp; <input type="radio" name="vMemb_GAP" value="0" <%=fcheck(fsqlboolean(vmemb_gap), 0)%>>NO<br>Select YES if this manager or facilitator can administer the GAP service.</td>
      </tr>
      <% Else %> 
      <input type="hidden" name="vMemb_GAP"     value="<%=fSqlBoolean(vMemb_GAP)%>">
      <% End If %> 


      <tr>

        <td align="center" width="100%" valign="top" colspan="2">
        <br>
       
          <% If Len(vNext) > 0 Then %>
            <%=f5%><input onclick="location.href='<%=fjUnQuote(vNext)%>'" type="button" value="<%=bReturn%>" name="bReturn" id="bReturn"class="button085"><%=f5%>
          <% End If %>
 
          <% If svCustAcctId <> fDefault(vMemb_AcctId, svCustAcctId) Then %>
             <h5><br><br><!--webbot bot='PurpleText' PREVIEW='Learner Profiles accessed from another Account are Read Only.'--><%=fPhra(001265)%><br><!--webbot bot='PurpleText' PREVIEW='You cannot Update or Delete this Learner's Profile'--><%=fPhra(001266)%>.<br></h5>
          <% Else %>     


          <%   If ((svMembLevel = 4 And vCust_DeleteLearners And vCust_MaxUsers = 0) Or svMembManager Or svMembLevel = 5) Then %><%=f5%>
                 <%=f5%><input onclick="jconfirm('User<%=fGroup%>.asp?vDelete=<%=vMemb_No%>&vNext=<%=Server.UrlEncode(vNext)%>','<!--webbot bot='PurpleText' PREVIEW='Ok to delete?'--><%=fPhra(000199)%>')" type="button" value="<%=bDelete%>" name="bDelete" class="button085"><%=f5%>
          <%   End If %>  


          <%   If ((svMembLevel < 5 And vCust_InsertLearners And vCust_UpdateLearners) Or svMembManager Or svMembLevel = 5) Then %> 
          <%     If svMembLevel > 3 Or vCust_MaxUsers = 0 Or (vCust_MaxUsers > 0 And (vCust_MaxUsers - fAllMembCount) > -3) Then %>
                   <%=f5%><input type="submit" value="<%=bUpdate%>" name="bUpdate" class="button085">
          <%     End If %> 
          <%   End If %>  

          <% End If %>

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

