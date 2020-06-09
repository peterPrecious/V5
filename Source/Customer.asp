<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/QueryString.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->
<!--#include virtual = "V5/Inc/CustomCertRoutines.asp"-->


<% 
  Dim vAddCustomerId, vEditCustId, vCloneCustId, vFunction, aProgs, aProg, vProgram, vNext, vMsg

  vFunction = Request("vFunction")

  '...these functions come from this form - note this record exists as a placeholder so need to properly insert it with all the suppporting functions
  If vFunction = "add" Then
    sExtractCust
'   sInsertCust
    sUpdateCust
    Session("CustLevel") = vCust_Level
    If Len(vNext) > 0 Then Response.Redirect vNext
    vMsg = "Customer was Inserted Successfully"

  ElseIf vFunction = "edit" Then
    sExtractCust
    sUpdateCust
    Session("CustLevel") = vCust_Level
    If Len(vNext) > 0 Then Response.Redirect vNext
    vMsg = "Customer was Updated Successfully"

  ElseIf Len(Request("vDelCustAcctId")) >= 4 Then 
    vCust_Id = Request("vDelCustId")
    vCust_AcctId = Request("vDelCustAcctId")
    sDeleteCust
    Response.Redirect "Customers.asp"

  ElseIf Len(Request("vDelCustId")) >= 8 Then 
    vCust_Id = Request("vDelCustId")
    sDeleteLinkedCust
    Response.Redirect "Customers.asp"

  End If  
   
  '...these functions come from Customers.asp
  If Request("vAddCustId").Count > 0 Then 
    vCust_Id = Ucase(Request("vAddCustId"))
    If Len(vCust_Id) <> 8 Then
      Response.Redirect "Error.asp?vErr=Invalid Customer Id.  Must be XXXXAAAA!"
    End If        
    vFunction = "add"

  ElseIf Request("vCloneCustId").Count > 0 Then 
    vCust_Id = Ucase(Request("vCustId"))
    vFunction = "add"

  ElseIf Len(Request.QueryString("vEditCustId")) >= 8 Then 
    vCust_Id = Request.QueryString("vEditCustId")
    vFunction = "edit"

  End If

  '...get the values (even if trying to add)
  sGetCust (vCust_Id)

  '...if Cloning use new CustId
  If Request("vCloneCustId").Count > 0 Then 
    vCust_Id     = Ucase(Request("vCloneCustId"))
    vCust_AcctId = ""
    vCust_Placeholder = "1"  '...this tells  sUpdateCust to create Internals / Repository
  End If

  '...update the Session variable
  If Len(vCust_ReturnUrl) > 0 Then  
    Session("CustReturnUrl")  = vCust_ReturnUrl
    svCustReturnUrl           = vCust_ReturnUrl
  End If

  If fNoValue(vCust_EcomGroupLicense) Then vCust_EcomGroupLicense = 3
  If fNoValue(vCust_EcomGroupSeat)    Then vCust_EcomGroupSeat    = .2
  If fNoValue(vCust_EcomGroup2Rates)  Then vCust_EcomGroup2Rates  = "5|20~10|30~25|40~50|50~200|60"

  If fNoValue(vCust_Tab1)             Then vCust_Tab1             = 1
  If fNoValue(vCust_Tab2)             Then vCust_Tab2             = 1
  If fNoValue(vCust_Tab3)             Then vCust_Tab3             = 1
  If fNoValue(vCust_Tab4)             Then vCust_Tab4             = 0
  If fNoValue(vCust_Tab5)             Then vCust_Tab5             = 1
  If fNoValue(vCust_Tab6)             Then vCust_Tab6             = 0
  If fNoValue(vCust_InfoEditProfile)  Then vCust_InfoEditProfile  = 1

  If fNoValue(vCust_Active)           Then vCust_Active           = 1

  If fNoValue(vCust_InsertLearners)   Then vCust_InsertLearners   = 1
  If fNoValue(vCust_UpdateLearners)   Then vCust_UpdateLearners   = 1
  If fNoValue(vCust_DeleteLearners)   Then vCust_DeleteLearners   = 1
  If fNoValue(vCust_ResetLearners)    Then vCust_ResetLearners    = 1
%>

<html>

<head>
  <title>Customer</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <script>
    $(function(){
      $(".bgGrey").hide();   
    });

    function Validate(theForm) {

      if (theForm.vCust_Agent.selectedIndex < 0) {
        alert("Please select one of the \"Agent\" options.");
        theForm.vCust_Agent.focus();
        return (false);
      }

      if (theForm.vCust_MaxSponsor.value == "") {
        alert("Please enter a value for the \"Max Sponsor\" field.");
        theForm.vCust_MaxSponsor.focus();
        return (false);
      }

      if (theForm.vCust_MaxSponsor.value.length < 1) {
        alert("Please enter at least 1 characters in the \"Max Sponsor\" field.");
        theForm.vCust_MaxSponsor.focus();
        return (false);
      }

      if (theForm.vCust_MaxSponsor.value.length > 2) {
        alert("Please enter at most 2 characters in the \"Max Sponsor\" field.");
        theForm.vCust_MaxSponsor.focus();
        return (false);
      }

      var checkOK = "0123456789-";
      var checkStr = theForm.vCust_MaxSponsor.value;
      var allValid = true;
      var validGroups = true;
      var decPoints = 0;
      var allNum = "";
      for (i = 0; i < checkStr.length; i++) {
        ch = checkStr.charAt(i);
        for (j = 0; j < checkOK.length; j++)
          if (ch == checkOK.charAt(j))
            break;
        if (j == checkOK.length) {
          allValid = false;
          break;
        }
        allNum += ch;
      }
      if (!allValid) {
        alert("Please enter only digit characters in the \"Max Sponsor\" field.");
        theForm.vCust_MaxSponsor.focus();
        return (false);
      }

      var chkVal = allNum;
      var prsVal = parseInt(allNum);
      if (chkVal != "" && !(prsVal >= 0 && prsVal <= 12)) {
        alert("Please enter a value greater than or equal to \"0\" and less than or equal to \"12\" in the \"Max Sponsor\" field.");
        theForm.vCust_MaxSponsor.focus();
        return (false);
      }

      if (theForm.vCust_EcomSplit.value.length > 3) {
        alert("Please enter at most 3 characters in the \"Ecommerce Customer Split\" field.");
        theForm.vCust_EcomSplit.focus();
        return (false);
      }

      var checkOK = "0123456789-";
      var checkStr = theForm.vCust_EcomSplit.value;
      var allValid = true;
      var validGroups = true;
      var decPoints = 0;
      var allNum = "";
      for (i = 0; i < checkStr.length; i++) {
        ch = checkStr.charAt(i);
        for (j = 0; j < checkOK.length; j++)
          if (ch == checkOK.charAt(j))
            break;
        if (j == checkOK.length) {
          allValid = false;
          break;
        }
        allNum += ch;
      }
      if (!allValid) {
        alert("Please enter only digit characters in the \"Ecommerce Customer Split\" field.");
        theForm.vCust_EcomSplit.focus();
        return (false);
      }

      var chkVal = allNum;
      var prsVal = parseInt(allNum);
      if (chkVal != "" && !(prsVal >= 0 && prsVal <= 100)) {
        alert("Please enter a value greater than or equal to \"0\" and less than or equal to \"100\" in the \"Ecommerce Customer Split\" field.");
        theForm.vCust_EcomSplit.focus();
        return (false);
      }

      if (theForm.vCust_AssessmentAttempts.value == "") {
        alert("Please enter a value for the \"Max Attempts\" field.");
        theForm.vCust_AssessmentAttempts.focus();
        return (false);
      }

      if (theForm.vCust_AssessmentAttempts.value.length < 1) {
        alert("Please enter at least 1 characters in the \"Max Attempts\" field.");
        theForm.vCust_AssessmentAttempts.focus();
        return (false);
      }

      if (theForm.vCust_AssessmentAttempts.value.length > 2) {
        alert("Please enter at most 2 characters in the \"Max Attempts\" field.");
        theForm.vCust_AssessmentAttempts.focus();
        return (false);
      }

      var checkOK = "0123456789-";
      var checkStr = theForm.vCust_AssessmentAttempts.value;
      var allValid = true;
      var validGroups = true;
      var decPoints = 0;
      var allNum = "";
      for (i = 0; i < checkStr.length; i++) {
        ch = checkStr.charAt(i);
        for (j = 0; j < checkOK.length; j++)
          if (ch == checkOK.charAt(j))
            break;
        if (j == checkOK.length) {
          allValid = false;
          break;
        }
        allNum += ch;
      }
      if (!allValid) {
        alert("Please enter only digit characters in the \"Max Attempts\" field.");
        theForm.vCust_AssessmentAttempts.focus();
        return (false);
      }

      var chkVal = allNum;
      var prsVal = parseInt(allNum);
      if (chkVal != "" && !(prsVal >= 0 && prsVal <= 99)) {
        alert("Please enter a value greater than or equal to \"0\" and less than or equal to \"99\" in the \"Max Attempts\" field.");
        theForm.vCust_AssessmentAttempts.focus();
        return (false);
      }

      var checkOK = "0123456789-.";
      var checkStr = theForm.vCust_AssessmentScore.value;
      var allValid = true;
      var validGroups = true;
      var decPoints = 0;
      var allNum = "";
      for (i = 0; i < checkStr.length; i++) {
        ch = checkStr.charAt(i);
        for (j = 0; j < checkOK.length; j++)
          if (ch == checkOK.charAt(j))
            break;
        if (j == checkOK.length) {
          allValid = false;
          break;
        }
        if (ch == ".") {
          allNum += ".";
          decPoints++;
        }
        else
          allNum += ch;
      }
      if (!allValid) {
        alert("Please enter only digit characters in the \"Assessment Score\" field.");
        theForm.vCust_AssessmentScore.focus();
        return (false);
      }

      if (decPoints > 1 || !validGroups) {
        alert("Please enter a valid number in the \"vCust_AssessmentScore\" field.");
        theForm.vCust_AssessmentScore.focus();
        return (false);
      }

      var chkVal = allNum;
      var prsVal = parseFloat(allNum);
      if (chkVal != "" && !(prsVal >= 0 && prsVal <= 1)) {
        alert("Please enter a value greater than or equal to \"0\" and less than or equal to \"1\" in the \"Assessment Score\" field.");
        theForm.vCust_AssessmentScore.focus();
        return (false);
      }
      return (true);
    }
  </script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body onload="parent.frames.tabs.location.href='tabslive.asp?vTab=9'">

  <% 
  	Server.Execute vShellHi
  %>

  <form method="POST" action="Customer.asp" target="_self" onsubmit="return Validate(this)">

    <input type="hidden" name="vFunction" value="<%=vFunction%>">
    <input type="hidden" name="vCust_Id" value="<%=vCust_Id%>">
    <input type="hidden" name="vCust_Placeholder" value="<%=vCust_Placeholder%>">
    <input type="hidden" name="vNext" value="<%=vNext%>">

    <table class="table">
      <tr>
        <td colspan="2" style="text-align: center">
          <h2><a <%=fstatx%> name="Top" class="c2">Customer Profile</a></h2>
          <% If Len(vMsg) > 0 Then %><h5><%=vMsg%> </h5>
          <% End If %>
          <h2>Select the Section(s) you wish to access.</h2>
          <h3>
            <a onclick="toggle('Div_Notes');" <%=fstatx%> href="#Notes">Basics</a> | 
            <a onclick="toggle('Div_Basics');" <%=fstatx%> href="#Basics">Core</a> | 
            <a onclick="toggle('Div_FeatureSet');" <%=fstatx%> href="#FeatureSet">Advanced</a> | 
            <span class="bgGrey">
            <a onclick="toggle('Div_Programs');" <%=fstatx%> href="#Programs">VuBuild</a> | 
            </span>
            <a onclick="toggle('Div_Ecommerce');" <%=fstatx%> href="#Ecommerce">Ecommerce</a> | 

            <span class="bgGrey">
            <a onclick="toggle('Div_Certificates');" <%=fstatx%> href="#Certificates">Certificates</a> | 
            </span>


            <a onclick="toggle('Div_Tabs');" <%=fstatx%> href="#Tabs">Tabs Setup</a> | 
            <a <%=fstatx%> href="#Bottom">Bottom</a>
            <br><br>
            <a onclick="openDivs('Div_')" href="#">Show All</a> <%=f5%> <a onclick="hideDivs('Div_')" href="#">Hide All</a><%=f10%><%=f10%>
            <a onclick="$('.bgGrey').show()" href="#">Show Old</a> <%=f5%> <a onclick="$('.bgGrey').hide()" href="#">Hide Old</a><br>
          </h3>
        </td>
      </tr>
      <tr>
        <th>Customer Id : </th>
        <td><%=vCust_Id%></td>
      </tr>
    </table>

    <div id="Div_Notes" class="div">
      <table class="table">
        <tr>
          <td style="text-align: center" colspan="2">
            <h2><a <%=fstatx%> name="Notes" class="c2">Customer Care</a></h2>
            <h3 style="text-align: left">Enter the basic setup features plus any notes specific to this Account. Note: these fields can also be maintained using the Customer Care System.</h3>
            <p style="text-align: right">
              <br>
              <a href="#Top">
                <img border="0" src="../Images/Icons/Ascending.gif" width="9" height="7"></a><a href="#Bottom"><img border="0" src="../Images/Icons/Descending.gif" width="9" height="7"></a>
            </p>
          </td>
        </tr>
        <tr>
          <th>Title : </th>
          <td>
            <input type="text" size="46" name="vCust_Title" value="<%=vCust_Title%>">
          </td>
        </tr>
        <tr>
          <th>Logo : </th>
          <td>
            <input type="text" size="80" name="vCust_Banner" value="<%=vCust_Banner_Original%>"><br>Name of the logo that appears on the banner, ie cfib.jpg or cfib_en.jpg.&nbsp; Must be 150 x 50 pixels with white background.&nbsp; Before going live, FTP the image to the Logos folder. Note to appear in flash, this must be a &quot;JPG&quot;. If there are different values for different languages, then separate them with pipes in EN|FR|ES order.
          </td>
        </tr>
        <tr>
          <th>Site URL : </th>
          <td>
            <input type="text" size="80" name="vCust_URL" value="<%=vCust_URL_Original%>"><br>Where to go when learner click on the Customer logo - typically the customer&#39;s home page.&nbsp; Note: do not precede with &quot;//&quot; and do NOT use Vubiz.com. If there are different URLS for different languages, then separate them with pipes in EN|FR|ES order.
          </td>
        </tr>
        <tr>
          <th>Start URL : </th>
          <td>
            <input type="text" size="80" name="vCust_StartUrl" value="<%=vCust_StartUrl_Original%>"><br>This is where the customer goes to sign in or enroll into Vubiz. ie /ChAccess/CAST/default_EN.asp. Note: do not precede with &quot;//&quot;. If there are different URLS for different languages, then separate them with pipes in EN|FR|ES order.
          </td>
        </tr>
        <tr>
          <th>Return URL : </th>
          <td>
            <input type="text" size="80" name="vCust_ReturnUrl" value="<%=vCust_ReturnUrl_Original%>"><br>Go to this URL on Sign Off. Leave empty or enter the full address (ie &quot;//cnn.com/training.php&quot;) or enter simply &quot;close&quot; if the browser window should close on Sign Off (used when clients open a separate window for Vubiz training). NOTE: this value will over-ride any value sent via a Landing Page. If you have different return URLs for different languages, then separate them with pipes in EN|FR|ES order.
          </td>
        </tr>
        <tr>
          <th>Email : </th>
          <td>
            <input type="text" size="44" name="vCust_Email" value="<%=vCust_Email%>">
            <br>Email address that appears on assorted pages.
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Application : </th>
          <td>
            <textarea name="vCust_Note1" <%=fmaxlength(8000)%>><%=vCust_Note1%></textarea>
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Content : </th>
          <td>
            <textarea name="vCust_Note2" <%=fmaxlength(8000)%>><%=vCust_Note2%></textarea>
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Assessment : </th>
          <td>
            <textarea name="vCust_Note3" <%=fmaxlength(8000)%>><%=vCust_Note3%></textarea>
          </td>
        </tr>
        <tr class="bgGrey">
          <th>LMS / CMS : </th>
          <td>
            <textarea name="vCust_Note4" <%=fmaxlength(8000)%>><%=vCust_Note4%></textarea>
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Other : </th>
          <td>
            <textarea name="vCust_Note5" <%=fmaxlength(8000)%>><%=vCust_Note5%></textarea>
          </td>
        </tr>
      </table>
    </div>

    <div id="Div_Basics" class="div">
      <table class="table">
        <tr>
          <td colspan="2" style="text-align: center">
            <h2><a <%=fstatx%> name="Basics" class="c2">Core Feature Set</a></h2>
            <h3 style="text-align: left">The following fields are mandatory to setup a Customer record.</h3>
            <p style="text-align: right">
              <br>
              <a href="#Top">
                <img border="0" src="../Images/Icons/Ascending.gif" width="9" height="7"></a><a href="#Bottom"><img border="0" src="../Images/Icons/Descending.gif" width="9" height="7"></a>
            </p>
          </td>
        </tr>
        <% If fNoValue(vCust_AcctId) Then %>
        <tr>
          <th>Account Id :</th>
          <td>
            <input type="text" name="vCust_AcctId" size="5" value="<%=fDefault(vCust_AcctId, Mid(vCust_Id, 5))%>" maxlength="6">
            <br>This field defaults to the right 4+ numbers of Customer Id.&nbsp; If you want to <b>ATTACH</b> this Customer to another thereby sharing a common set of learners, then enter the right 4+ numbers of the other <b>Master</b> Customer Id. <font color="#FF0000">Note, once this record is saved the Account Id <b>CANNOT</b> be changed.</font>
          </td>
        </tr>
        <% Else %>
        <input type="hidden" name="vCust_AcctId" value="<%=vCust_AcctId%>">
        <tr>
          <th>Account Id :</th>
          <td><%=vCust_AcctId%></td>
        </tr>
        <% End If %>
        <tr>
          <th>Active : </th>
          <td>
            <input type="radio" name="vCust_Active" value="1" <%=fcheck(fsqlboolean(vcust_active), 1)%>>Yes&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_Active" value="0" <%=fcheck(fsqlboolean(vcust_active), 0)%>>No
          </td>
        </tr>
        <tr>
          <th>Date Added :</th>
          <td><%=fFormatDate(fDefault(vCust_Added, Now))%></td>
        </tr>
        <tr>
          <th>Date Modified :</th>
          <td><%=fFormatDate(vCust_Modified)%></td>
        </tr>


        <tr>
          <th>Account Type :</th>
          <td>

            <span class="bgGrey">
            <input type="radio" name="vCust_Level" class="custLevel" value="1" <%=fcheck(vcust_level, 1)%>>1 - Content Only (this just launches CDs etc). Set &quot;Auto&quot; to Yes.<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="text" name="vCust_ContentLaunch" size="43" value="<%=vCust_ContentLaunch%>"><br>
            </span>

            <input type="radio" name="vCust_Level" class="custLevel" value="2" <%=fcheck(vcust_level, 2)%>>2 - Channel<br>

            <div style="margin: 0 20px;">

              <span class="bgYellow">
                <input type="checkbox" name="vCust_CatalogueMaster" id="vCust_CatalogueMaster" value="1" <%=fcheck(fsqlboolean(vCust_CatalogueMaster), 1)%>><%=svCustId%> contains the <b>master</b> catalogue for <%=Left(svCustId, 4)%> siblings.<br>
                <input type="checkbox" name="vCust_CatalogueSibling" id="vCust_CatalogueSibling" value="1" <%=fcheck(fsqlboolean(vCust_CatalogueSibling), 1)%>>Is a <b>sibling</b> that auto inherits the <%=Left(svCustId, 4)%> master catalogue.<br>
              </span>

              <input type="checkbox" name="vCust_ChannelNop" id="vCust_ChannelNop" value="1" <%=fcheck(fsqlboolean(vCust_ChannelNop), 1)%>>Is NOP (Id/Username unique for all NOP accounts)<br>
              <input type="checkbox" name="vCust_ChannelV8" id="vCust_ChannelV8" value="1" <%=fcheck(fsqlboolean(vCust_ChannelV8), 1)%>>Is V8<br>
              <div style="margin: 0 20px;">
                <input type="checkbox" name="vCust_ChannelParent" id="vCust_ChannelParent" value="1" <%=fcheck(fsqlboolean(vCust_ChannelParent), 1)%>>Is G2 Parent<br>
                <input type="checkbox" name="vCust_ChannelReportsTo" id="vCust_ChannelReportsTo" value="1" <%=fcheck(fsqlboolean(vCust_ChannelReportsTo), 1)%>>Use V8 Uploads with ReportsTo<br>
                <input type="checkbox" name="vCust_ChannelGuests" id="vCust_ChannelGuests" value="1" <%=fcheck(fsqlboolean(vCust_ChannelGuests), 1)%>>Use V8 Guest subsystem<br>
              </div>
            </div>

            <input type="radio" name="vCust_Level" class="custLevel" value="4" <%=fcheck(vcust_level, 4)%>>4 - Corporate<br>
            <span class="bgGrey">
            <input type="radio" name="vCust_Level" class="custLevel" value="7" <%=fcheck(vcust_level, 7)%>>7 - vuBuild Account (for authentication)          
            </span>
            <br>
            <span class="red">Note: once set, do NOT change the Account Type since Corporate accounts create a Repository file set.</span>
          </td>
        </tr>

        <script>
            $(".custLevel").on("click", function() {

              if ( $(".custLevel:checked").val() == "2")
              {
                $("#vCust_ChannelParent").prop("checked", true);
                $("#vCust_ChannelV8").prop("checked", true);
                $("#vCust_ChannelReportsTo").prop("checked" , true);
                $("#vCust_ChannelGuests").prop("checked", true);
              } else {
                $("#vCust_ChannelParent").prop("checked", false);
                $("#vCust_ChannelV8").prop("checked", false);
                $("#vCust_ChannelReportsTo").prop("checked", false);
                $("#vCust_ChannelGuests").prop("checked", false);
              };          
  
            });
        </script>

        <tr>
          <th>Parent Id :</th>
          <td>
            <input type="text" name="vCust_ParentId" size="4" value="<%=vCust_ParentId%>" maxlength="6"><br>If filled (ie 1234), this is the Parent Account ID that cloned this Customer record.&nbsp; It is used by the Custom and Self Serve Email Alert Systems<font color="#FF0000">. This field should <b>NOT</b> be modified.</font>
          </td>
        </tr>
        <tr>
          <th>Languages :</th>
          <td>
            <input type="checkbox" name="vCust_Lang" value="EN" <%=fchecks(vcust_lang, "en")%>>EN<br>
            <input type="checkbox" name="vCust_Lang" value="FR" <%=fchecks(vcust_lang, "fr")%>>FR<br>
            <input type="checkbox" name="vCust_Lang" value="ES" <%=fchecks(vcust_lang, "es")%>>ES<br>
            <span style="background-color: grey" class="bgGrey">
              <input type="checkbox" name="vCust_Lang" value="PT" <%=fchecks(vcust_lang, "pt")%>>PT<br></span>
            Select the languages offered in Channel or Corporate unless using the V8 shell, in which case 
            <span class="bgYellow">select just the "primary language" - typically only needed if multi-lingual and preferred language is NOT EN.</span>
          </td>
        </tr>
        <tr>
          <th>Partner (Agent) :</th>
          <td>
            <!--webbot bot="Validation" s-display-name="Agent" b-value-required="TRUE" -->
            <select size="1" name="vCust_Agent">
              <option <%=fselect("VUBZ", vcust_agent)%> value="VUBZ">VUBZ</option>
              <option <%=fselect("CCHS", vcust_agent)%> value="CCHS">CCHS</option>
              <option <%=fselect("IAPA", vcust_agent)%> value="IAPA">IAPA</option>
              <option <%=fselect("ERGP", vcust_agent)%> value="ERGP">ERGP</option>
              <option <%=fselect("WSPS", vcust_agent)%> value="WSPS">WSPS</option>
              <option <%=fselect("DUAL", vcust_agent)%> value="DUAL">DUAL (CCHS+WSPS)</option>
            </select><br>This is used in ecommerce related reports.
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Max Learners :</th>
          <td>
            <input type="text" size="3" name="vCust_MaxUsers" value="<%=vCust_MaxUsers%>">
            <br>Generally leave empty for no restrictions.&nbsp; Used for Group 1 Ecommerce sites and Corporate sites.
          </td>
        </tr>
        <tr class="bgGrey">
          <th colspan="2">
            <p><br>This next 4 options override normal rights and do NOT apply to Super Managers nor Administrators<br></p>
          </th>
        </tr>
        <tr class="bgGrey">
          <th>Ok to Insert Learners :</th>
          <td>
            <input type="radio" name="vCust_InsertLearners" value="1" <%=fcheck(fsqlboolean(vcust_insertlearners), 1)%>>Yes&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_InsertLearners" value="0" <%=fcheck(fsqlboolean(vcust_insertlearners), 0)%>>No
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Ok to Update Learners :</th>
          <td>
            <input type="radio" name="vCust_UpdateLearners" value="1" <%=fcheck(fsqlboolean(vcust_updatelearners), 1)%>>Yes&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_UpdateLearners" value="0" <%=fcheck(fsqlboolean(vcust_updatelearners), 0)%>>No
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Ok to Delete Learners :</th>
          <td>
            <input type="radio" name="vCust_DeleteLearners" value="1" <%=fcheck(fsqlboolean(vcust_deletelearners), 1)%>>Yes&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_DeleteLearners" value="0" <%=fcheck(fsqlboolean(vcust_deletelearners), 0)%>>No
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Ok to Reset Scores :</th>
          <td>
            <input type="radio" name="vCust_ResetLearners" value="1" <%=fcheck(fsqlboolean(vcust_resetlearners), 1)%>>Yes&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_ResetLearners" value="0" <%=fcheck(fsqlboolean(vcust_resetlearners), 0)%>>No
          </td>
        </tr>
      </table>
    </div>

    <div id="Div_FeatureSet" class="div">
      <table class="table">
        <tr>
          <td style="text-align: center" colspan="2">
            <h2><a <%=fstatx%> name="FeatureSet" class="c2">Advanced Feature Set</a></h2>
            <h3 style="text-align: left">The service level determines what features are displayed.&nbsp; </h3>
            <p style="text-align: right">
              <br>
              <a href="#Top">
                <img border="0" src="../Images/Icons/Ascending.gif" width="9" height="7"></a>
              <a href="#Bottom">
                <img border="0" src="../Images/Icons/Descending.gif" width="9" height="7"></a>
            </p>
          </td>
        </tr>
        <tr>
          <th>Completion Service&nbsp; ?<br>&nbsp; </th>
          <td>
            <input type="radio" name="vCust_Completion" value="0" <%=fcheck(fsqlboolean(vcust_completion), 0)%>>No (default)<br>
            <input type="radio" name="vCust_Completion" value="1" <%=fcheck(fsqlboolean(vcust_completion), 1)%>>Yes<br />
            This is for clients with multiple locations who need their learner&#39;s Completion Status &quot;rolled up&quot; to the National/Regional and Location level.
          </td>
        </tr>
        <tr>
          <th>Ecommerce Profile :</th>
          <td>
            <input type="checkbox" name="vCust_EcomSeller" value="1" <%=fcheck(fsqlboolean(vcust_ecomseller), 1)%>>Seller (typically most channels but can be like SBMC, etc)<br>
            <input type="checkbox" name="vCust_EcomOwner" value="1" <%=fcheck(fsqlboolean(vcust_ecomowner), 1)%>>Owner (typically a channel but can be an Owner only)
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Authoring Rights ?</th>
          <td>
            <input type="radio" name="vCust_Auth" value="1" <%=fcheck(fsqlboolean(vcust_auth), 1)%>>Yes&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_Auth" value="0" <%=fcheck(fsqlboolean(vcust_auth), 0)%>>No<br>Select Yes if this account is authorized to use VuBuild and has been designated as a Level 7 Account.&nbsp; Remember to also enable the appropriate members who can author content.
          </td>
        </tr>
        <tr>
          <td colspan="2" style="text-align: center">
            <br>The next fields determine how the site handles/supports its learners.<br></td>
        </tr>
        <tr>
          <th>Learners Need Passwords ?</th>
          <td>
            <input type="radio" name="vCust_Pwd" value="1" <%=fcheck(fsqlboolean(vcust_pwd), 1)%>>Yes&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_Pwd" value="0" <%=fcheck(fsqlboolean(vcust_pwd), 0)%>>No<br>Default No, used on for corporate landing pages, not channels. If Yes, then each learner must enter with a valid password.
          </td>
        </tr>
        <tr>
          <th>SSO (Auto Enrol) ?</th>
          <td>
            <input type="radio" name="vCust_Auto" value="1" <%=fcheck(fsqlboolean(vcust_auto), 1)%>>Yes&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_Auto" value="0" <%=fcheck(fsqlboolean(vcust_auto), 0)%>>No
            <br>Allows SSO for learners only when authentication occurs behind the customer's firewall (mandatory for Level 1 Customers)
            <br>Set to "No" for V8 where the presence of a custGuid in the call invokes SSO
          </td>
        </tr>
        <tr>
          <th>Max Sponsor Learners :</th>
          <td>
            <!--webbot bot="Validation" s-display-name="Max Sponsor" s-data-type="Integer" s-number-separators="x" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="2" s-validation-constraint="Greater than or equal to" s-validation-value="0" s-validation-constraint="Less than or equal to" s-validation-value="12" -->
            <input type="text" name="vCust_MaxSponsor" size="2" value="<%=fDefault(vCust_MaxSponsor, 0)%>" maxlength="2"><br>If applicable, enter the number of learners each learner can sponsor.&nbsp; Leave 0 if none else typically 3 (max 12).&nbsp; If &gt; 0 then a Sponsor link appears on the Info Page for the original member (not the sponsored member).
          </td>
        </tr>
        <input type="hidden" name="vCust_ResetStatus" value="0">
        <tr>
          <th>Issue Passwords Online ?</th>
          <td>
            <input type="radio" name="vCust_IssueIds" value="1" <%=fcheck(fsqlboolean(vcust_issueids), 1)%>>Yes&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_IssueIds" value="0" <%=fcheck(fsqlboolean(vcust_issueids), 0)%>>No<br>Allows Issuing 9 character Passwords (note: it is better to use the &quot;add learners&quot; feature).
          </td>
        </tr>
        <tr>
          <th>Issue Passwords by&nbsp;&nbsp;
            <br>Email Template :</th>
          <td>
            <input type="text" size="11" name="vCust_IssueIdsTemplate" value="<%=vCust_IssueIdsTemplate%>" maxlength="5">If feature is required, enter the Email Template Id that contains the email response message (default message is E0000).
            <br>Leave field empty if feature is not required.<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_IssueIdsMemo" value="1" <%=fcheck(fsqlboolean(vcust_issueidsmemo), 1)%>>Yes&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_IssueIdsMemo" value="0" <%=fcheck(fsqlboolean(vcust_issueidsmemo), 0)%>>No<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; Include a memo field on the input?
          </td>
        </tr>
        <tr>
          <th>Activate Ids ?</th>
          <td>
            <input type="radio" name="vCust_ActivateIds" value="1" <%=fcheck(fsqlboolean(vcust_activateids), 1)%>>Yes&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_ActivateIds" value="0" <%=fcheck(fsqlboolean(vcust_activateids), 0)%>>No&nbsp;&nbsp;&nbsp;
            <br>Allows Activating Customer Ids (better to add/edit members).
          </td>
        </tr>
        <tr>
          <th>Ids Size :</th>
          <td>
            <input type="text" size="5" name="vCust_IdsSize" value="<%=vCust_IdsSize%>"><br>Useful for Auto (like CFIB).&nbsp; If not zero then Passwords must be this length.
          </td>
        </tr>
        <tr>
          <th>Memo :</th>
          <td>
            <textarea name="vCust_Desc"><%=vCust_Desc%></textarea><br>A free field, for internal usage only - do not use as a comment field.
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Expires :</th>
          <td>
            <input type="text" size="27" name="vCust_Expires" value="<%=fFormatDate(vCust_Expires)%>"><br>Normally leave empty if no limit else specify a date (ie Jan 15, 2007).&nbsp; Typically assigned for Group 1 Accounts and useful for demos.
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Seed Log File :</th>
          <td>
            <textarea name="vCust_SeedLogs"><%=vCust_SeedLogs%></textarea><br>Enter a list of key programs, ie P1234EN P1234ES P5001XX P5001ES, etc which will generate a dummy Time Spent log entry for each Program listed ensuring that the Programs are listed in the Learner Report Card even if the learner has not accessed them.&nbsp; If you use XX for the Language then it will be replaced by the Language selected by the Learner.
          </td>
        </tr>
      </table>
    </div>

    <div class="bgGrey">
    <div id="Div_Programs" class="div">
      <table class="table">
        <tr>
          <td style="text-align: center" colspan="2">
            <h2>
            <br>
            <a <%=fstatx%> name="Programs" class="c2">VuBuild Content</a></h2>
            <h3 style="text-align: left">This field defines Programs that are used for VuBuild.&nbsp; Enter the full Program String such as &quot;P1176EN~0~0~1~365&quot;.</h3>
            <p style="text-align: right">
              <br>
              <a href="#Top">
                <img border="0" src="../Images/Icons/Ascending.gif" width="9" height="7"></a><a href="#Bottom"><img border="0" src="../Images/Icons/Descending.gif" width="9" height="7"></a>
            </p>
          </td>
        </tr>
        <tr>
          <th>Programs :</th>
          <td>
            <textarea name="vCust_Programs"><%=vCust_Programs%></textarea><br>
            <a <%=fstatx%> href="ContentEditAll.asp?vCust_Id=<%=vCust_Id%>&vCust_Programs=<%=vCust_Programs%>">Edit String</a> <% 
            If Len(vCust_Programs) > 0 Then 
            %>
            <br>Click to view programs:<br>
            <% 
              aProgs = Split(Trim(vCust_Programs), " ")
              For i = 0 to Ubound(aProgs)
                aProg          = Split(aProgs(i), "~")
                vProg_Id       = aProg(0)
            %> <a <%=fstatx%> target="_blank" href="Program.asp?vEditProgID=<%=vProg_ID%>&vHidden=n"><%=vProg_Id%></a> <%
              Next
            %>
            <br>Click to view program details:<br>
            <% 
              For i = 0 to Ubound(aProgs)
                aProg          = Split(aProgs(i), "~")
                vProgram       = aProgs(i)
                vProg_Id       = aProg(0)
            %> <a <%=fstatx%> href="javascript:programwindow('<%=vProgram%>')"><%=vProg_Id%></a> <%
              Next
            End If 
            %>
          </td>
        </tr>
      </table>
    </div>
    </div>


    <div id="Div_Ecommerce" class="div">
      <table class="table">
        <tr>
          <td style="text-align: center" colspan="2">
            <h2>
            <br>
            <a <%=fstatx%> name="Ecommerce" class="c2">Ecommerce Feature Set</a></h2>
            <h3 style="text-align: left">While online Content is defined in the Customer Catalogue, CD content, via the Program strings are entered for sale via ecommerce.&nbsp; If Products are offered via ecommerce then they use the standard Product table (which are configured outside the system.&nbsp; If the US and CA prices are zero then the programs appear in &quot;My Content&quot;, else they appear in &quot;More Content&quot;.&nbsp; If you put in a nominal value of $1 in both the CA and US prices,&nbsp; then they are given out free with ANY ecommerce purchase.&nbsp; If you put in a nominal value of $9999 in both the CA and US prices, then they will remain available but will not appear on &quot;More Content&quot;.</h3>
            <p style="text-align: right">
              <br>
              <a href="#Top">
                <img border="0" src="../Images/Icons/Ascending.gif" width="9" height="7"></a><a href="#Bottom"><img border="0" src="../Images/Icons/Descending.gif" width="9" height="7"></a>
            </p>
          </td>
        </tr>
        <tr>
          <th style="text-align: right" width="30%" nowrap>Ecom Currency :</th>
          <td align="left" width="69%">
            <input type="radio" name="vCust_EcomCurrency" value="CA" <%=fcheck(vcust_ecomcurrency, "ca")%>>CA&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_EcomCurrency" value="US" <%=fcheck(vcust_ecomcurrency, "us")%>>US&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <br>Select core currency and the other will be computed based on 1 US = <%=vCurrency%> CA dollars.<br>Note: currency value is set in &quot;/V5/Inc/Setup.asp&quot; file.
          </td>
        </tr>
        <tr>
          <th>Ecom Offerings :</th>
          <td>
            <input type="checkbox" name="vCust_ContentOnline" value="1" <%=fcheck(fsqlboolean(vcust_contentonline), 1)%>>Individual License
            
            <span class="bgGrey">
            (Channel or Corporate as defined by Service Type above)<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; These 3 fields are for Corporate sites only:<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="text" name="vCust_EcomCorpRate" size="3" value="<%=vCust_EcomCorpRate%>">Per Seat Fee in currency above, ie 99.00.<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="text" name="vCust_EcomCorpDuration" size="3" value="<%=vCust_EcomCorpDuration%>">No of days access, ie 90.<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="text" name="vCust_EcomCorpProgram" size="9" value="<%=vCust_EcomCorpProgram%>" maxlength="7">Program No (ie P9909EN)
            </span>
            <br>
            <span class="bgGrey">
            <input type="checkbox" name="vCust_ContentGroup" value="1" <%=fcheck(fsqlboolean(vcust_contentgroup),  1)%>>Group 1 Multi-Learner License (by Learner) --- [legacy]<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="text" name="vCust_EcomGroupLicense" size="3" value="<%=vCust_EcomGroupLicense%>">Annual License as Ratio of Individual Pricing (3.0 if left empty)<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input type="text" name="vCust_EcomGroupSeat" size="3" value="<%=vCust_EcomGroupSeat%>">Per Seat Fee as Ratio of Individual Pricing (0.2 if left empty)<br>
            </span>

            <input type="checkbox" name="vCust_ContentGroup2" value="1" <%=fcheck(fsqlboolean(vcust_contentgroup2), 1)%>>Group License (by Program)

            <div style="margin:5px 0 10px 20px">
            <input type="text" name="vCust_EcomGroup2Rates" size="54" value="<%=vCust_EcomGroup2Rates%>">
            <br> 
            Specifiy pricing which is a discount off Single Learner Pricing. Defaults to 5|25~10|45~25|55~50|65~200|75 meaning 5-9 get 25% off, 10-24 get 45% off, etc.           
            Note: Only allows 5 ranges with a maximum of 500 seats.<br>
            </div>

            Automatically Email Alert G2 Learners who are assigned new content?<br />
            <input type="radio" name="vCust_EcomG2alert" value="1" <%=fcheck(fsqlboolean(vcust_ecomg2alert), 1)%>>Yes&nbsp;&nbsp;
            <input type="radio" name="vCust_EcomG2alert" value="0" <%=fcheck(fsqlboolean(vcust_ecomg2alert), 0)%>>No

            <span class="bgYellow"><br>Email Alerts are only available if the Parent Customer has this flag enabled AND the appropriate Parent templates have been setup in the Alert System.</span>
            <br>
            <p class="c6">Note: One or more of the above License Type (Ecom Offerings) are mandatory if customer offers ecommerce and the ecommerce tab is set below.</p>
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Send Purchase Confirmation :</th>
          <td>
            <input type="radio" name="vCust_EcomConfirmation" value="1" <%=fcheck(fsqlboolean(vcust_ecomconfirmation), 1)%>>Yes
            <input type="radio" name="vCust_EcomConfirmation" value="0" <%=fcheck(fsqlboolean(vcust_ecomconfirmation), 0)%>>No
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Confirmation Email Message :</th>
          <td>
            <textarea style="height: 200px" name="vCust_EcomEmailBody" cols="73" <%=fmaxlength(8000)%>><%=vCust_EcomEmailBody%></textarea>
          </td>
        </tr>
        <tr>
          <td colspan="2" style="text-align: center">
            <h2>Ecommerce Contract &amp; Discounts</h2>
            <h3>Discounts apply only to Individual Sales not Group Sales!</h3>
            <p style="text-align: right">
              <br>
              <a href="#Top">
                <img border="0" src="../Images/Icons/Ascending.gif" width="9" height="7"></a><a href="#Bottom"><img border="0" src="../Images/Icons/Descending.gif" width="9" height="7"></a>
            </p>
          </td>
        </tr>
        <tr>
          <th>Ecommerce Revenue Split :</th>
          <td>
            <!--webbot bot="Validation" s-display-name="Ecommerce Customer Split" s-data-type="Integer" s-number-separators="x" i-maximum-length="3" s-validation-constraint="Greater than or equal to" s-validation-value="0" s-validation-constraint="Less than or equal to" s-validation-value="100" -->
            <input type="text" name="vCust_EcomSplit" size="3" value="<%=vCust_EcomSplit%>" maxlength="3">% Customer receives from Ecom sale (ie 30%)
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Discount Options :</th>
          <td>
            <input type="radio" value="0" name="vCust_EcomDiscOptions" <%=fcheck(vcust_ecomdiscoptions, 0)%>>No discounts<br>
            <input type="radio" value="1" name="vCust_EcomDiscOptions" <%=fcheck(vcust_ecomdiscoptions, 1)%>>Basic discount only<br>
            <input type="radio" value="2" name="vCust_EcomDiscOptions" <%=fcheck(vcust_ecomdiscoptions, 2)%>>Repurchase discount only<br>
            <input type="radio" value="3" name="vCust_EcomDiscOptions" <%=fcheck(vcust_ecomdiscoptions, 3)%>>Repurchase discount (if it applies) else Basic discount.<br>
            <input type="radio" value="4" name="vCust_EcomDiscOptions" <%=fcheck(vcust_ecomdiscoptions, 4)%>>Both discounts<br>
            <br>Note: selecting &quot;Both&quot; applies the basic discount first then the repurchase is a discount off the already discounted value.&nbsp; Ie a 20% basic discount on a $100 program costs $80 and a further 50% repurchase discount gives you a cost of $40.
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Restrict discount to programs :</th>
          <td>
            <input type="text" size="59" name="vCust_EcomDiscPrograms" value="<%=vCust_EcomDiscPrograms%>" maxlength="2000">
            <br>Enter programs, separated by spaces (ie P1002EN P10006EN) to which discounts apply. Leave empty if discount applies to all programs.
          </td>
        </tr>
        <tr class="bgGrey">
          <th>Apply a basic discount of :</th>
          <td>
            <input type="text" size="4" name="vCust_EcomDisc" value="<%=vCust_EcomDisc%>" maxlength="2">% (1-99, 0 if no discount)
          </td>
        </tr>
        <tr class="bgGrey">
          <th>if minimum $US order :</th>
          <td>
            <input type="text" size="6" name="vCust_EcomDiscMinUS" value="<%=vCust_EcomDiscMinUS%>">$US&nbsp; Only apply discount if order exceeds this $US amount.&nbsp; If no minimum order required, leave at zero.
          </td>
        </tr>
        <tr class="bgGrey">
          <th>or if minimum $CA order :</th>
          <td>
            <input type="text" size="6" name="vCust_EcomDiscMinCA" value="<%=vCust_EcomDiscMinCA%>">$CA&nbsp; Only apply discount if order exceeds this $CA amount.&nbsp; If no minimum order required, leave at zero.
            <br>
            <br>Note: If you enter a $US and a $CA value, then both must be true to apply the discount.
          </td>
        </tr>
        <tr class="bgGrey">
          <th>or if minimum no of&nbsp;&nbsp;&nbsp;
            <br>programs ordered :</th>
          <td>
            <input type="text" size="4" name="vCust_EcomDiscMinQty" value="<%=vCust_EcomDiscMinQty%>" tabindex="3"># programs, ie enter 2 if &quot;get 20% discount if you purchase 2 or more programs&quot;&nbsp; if no minimum quantity required, leave at zero
          </td>
        </tr>
        <tr class="bgGrey">
          <th>limit discount to :</th>
          <td>
            <input type="text" size="4" name="vCust_EcomDiscLimit" value="<%=vCust_EcomDiscLimit%>" tabindex="3">
            Enter no of programs (over any minimum required) that qualify for the discount.&nbsp; Set to zero if there are no restrictions (ie all ordered programs get the discount).&nbsp; Ie if this is a buy 1 get second at 50% then enter 1. Note: If you enter a number x, then the discount applies to the next x program&#39;s) that appears in the Ecom basket, ie the learner cannot determine what is the &quot;next&quot; program&#39;s).
          </td>
        </tr>
        <tr class="bgGrey">
          <th>excluding the original&nbsp;&nbsp;&nbsp;
            <br>programs ordered :</th>
          <td>
            <input type="checkbox" name="vCust_EcomDiscOriginal" value="1" <%=fcheck(fsqlboolean(vcust_ecomdiscoriginal), 1)%>>
            Check if discount applies to any minimum programs required to qualify, or leave unchecked if discount only applies to extra programs ordered.&nbsp; (ie tick if &quot;get 20% of 3 or more programs&quot;, but unchecked if &quot;buy 1 get 2nd at 20% off).<br></td>
        </tr>
        <tr class="bgGrey">
          <th>Apply repurchase discount of :</th>
          <td>
            <input type="text" size="4" name="vCust_EcomRepurDisc" value="<%=vCust_EcomRepurDisc%>" tabindex="3">%. Percentage applies to programs sold within time period below.
          </td>
        </tr>
        <tr class="bgGrey">
          <th>if repurchased within :</th>
          <td>
            <input type="text" size="4" name="vCust_EcomRepurPeriod" value="<%=vCust_EcomRepurPeriod%>" maxlength="3">days. No of days repurchase is valid: ie 30, 90, 365 (max 999).&nbsp; If no restrictions, leave as 0.
          </td>
        </tr>
        <tr class="bgGrey">
          <th>but just on the&nbsp;&nbsp;&nbsp;&nbsp;
            <br>program repurchased :</th>
          <td>
            <input type="checkbox" name="vCust_EcomRepurPrograms" value="1" <%=fcheck(fsqlboolean(vcust_ecomrepurprograms), 1)%>>Check if repurchase discount applies to all selected programs, or leave unchecked if discount only applies to a program previously ordered.<br></td>
        </tr>
      </table>
    </div>

    <div class="bgGrey">
    <div id="Div_Certificates" class="div">
      <table class="table">
        <tr>
          <td style="text-align: center" colspan="2">
            <h2><a <%=fstatx%> name="Certificates" class="c2">Certificates</a></h2>
            <h3 align="left">The Host/Customer fields determine which logos appear on the certificate. Enter &quot;n&quot; if no logo, leave blank to default to both host and Customer logos or enter a specific logo that is in the logos folder.&nbsp; Email Alerts can be sent to inform customers that a client (with an email address) completed a program/exam.</h3>
            <p style="text-align: right">
              <br>
              <a href="#Top">
                <img border="0" src="../Images/Icons/Ascending.gif" width="9" height="7"></a><a href="#Bottom"><img border="0" src="../Images/Icons/Descending.gif" width="9" height="7"></a>
            </p>
          </td>
        </tr>
        <tr>
          <th>Assessment Survey ?</th>
          <td>
            <input type="checkbox" name="vCust_Survey" value="1" <%=fcheck(fsqlboolean(vcust_survey), 1)%>>If checked, standard survey button appears in the program status (EN/FR).
          </td>
        </tr>
        <tr>
          <th>No Certificate ?</th>
          <td>
            <input type="checkbox" name="vCust_NoCert" value="1" <%=fcheck(fsqlboolean(vcust_nocert), 1)%>>If checked, ignore following fields.
          </td>
        </tr>
        <tr>
          <th>Legacy Custom Certificate ?</th>
          <td>
            <input type="checkbox" name="vCust_CustomCert" value="1" <%=fcheck(fsqlboolean(vcust_customcert), 1)%>>If checked, then &quot;certificate.asp&quot; is in /Repository/<%=svCustAcctId%>/Tools folder.&nbsp; Typically replaced by Custom Certificate below.
          </td>
        </tr>
        <tr>
          <th>Host Logo :</th>
          <td>
            <input type="text" size="31" name="vCust_CertLogoVubiz" value="<%=vCust_CertLogoVubiz%>">
            Enter &quot;n&quot; if no logo
          </td>
        </tr>
        <tr>
          <th>Customer Logo :</th>
          <td>
            <input type="text" size="31" name="vCust_CertLogoCust" value="<%=vCust_CertLogoCust%>">
            Enter &quot;n&quot; if no logo
          </td>
        </tr>
        <tr>
          <th>Email Alert :</th>
          <td>
            <input type="text" size="31" name="vCust_CertEmailAlert" value="<%=vCust_CertEmailAlert%>"><br>Leave empty if no alert sent. Enter a Email Address where alerts are sent if a Cert is issued. An English Email message list cert, learner and email address.
          </td>
        </tr>
        <!--- use for new certificate features -->
        <tr>
          <th align="left" colspan="2">
            <p>
              <br>The following are for the VuAssess system (ie Not for V5 Platform Tests and Exams) ...<br></p>
          </th>
        </tr>
        <tr>
          <th style="text-align: right" nowrap width="30%">Max Attempts :</th>
          <td>
            <!--webbot bot="Validation" s-display-name="Max Attempts" s-data-type="Integer" s-number-separators="x" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="2" s-validation-constraint="Greater than or equal to" s-validation-value="0" s-validation-constraint="Less than or equal to" s-validation-value="99" -->
            <input type="text" size="2" name="vCust_AssessmentAttempts" value="<%=fDefault(vCust_AssessmentAttempts, 0)%>" maxlength="2">
            If the max attempts for any VuAssess Test is other than the default of 3 (in which case leave as 0), specify the maximum no of attempts&nbsp; where 99 specifies no restrictions.
          </td>
        </tr>
        <tr>
          <th style="text-align: right" nowrap width="30%">Passing Score :</th>
          <td>
            <!--webbot bot="Validation" s-display-name="Assessment Score" s-data-type="Number" s-number-separators="x." s-validation-constraint="Greater than or equal to" s-validation-value="0" s-validation-constraint="Less than or equal to" s-validation-value="1" -->
            <input type="text" name="vCust_AssessmentScore" size="2" value="<%=fDefault(vCust_AssessmentScore, 0)%>">
            Score needed to display a certificate (must be between 0 and 1 - ie 70% is entered as .7).&nbsp; If you enter zero then 80% is assumed (.8).&nbsp; If you wish to display a certificate regardless of score then enter .01.
          </td>
        </tr>
      </table>
    </div>
    </div>



    <div id="Div_Tabs" class="div">
      <table class="table">

        <tr>
          <td colspan="2" style="text-align: center">
            <h2><a <%=fstatx%> name="Tabs" class="c2">Tabs Setup</h2>
            <h3 style="text-align:left">The following enables you to configure/hide tabs which would normally appear.</h3>

            <p style="text-align: right">
              <br>
              <a href="#Top"><img border="0" src="../Images/Icons/Ascending.gif" width="9" height="7"></a><a href="#Bottom"><img border="0" src="../Images/Icons/Descending.gif" width="9" height="7"></a>
            </p>
          </td>
        </tr>

        <tr class="bgGrey">
          <th>Header/Cluster Page Id :</th>
          <td>
            <input type="text" size="7" name="vCust_Cluster" value="<%=vCust_Cluster%>" maxlength="5">
            Defaults to C0001.&nbsp; This page will become part of the info page.
          </td>
        </tr>

        <tr>
          <th>&nbsp;Tab 1 :</th>
          <td>
            <input type="checkbox" name="vCust_Tab1" value="1" <%=fcheck(fsqlboolean(vcust_tab1), 1)%>>Info Page
            <br>&nbsp;&nbsp;&nbsp;&nbsp; Name :
            <input type="text" name="vCust_Tab1Name" size="60" value="<%=vCust_Tab1Name%>" maxlength="62"><br>&nbsp;&nbsp;&nbsp;&nbsp; Leave empty for default, else enter name as EN|FR|ES.&nbsp; <font color="#FF0000">Individual Tab names <br>&nbsp;&nbsp;&nbsp;&nbsp; cannot can exceed 20 chars in length</font>.<br>
            <br>&nbsp;&nbsp;&nbsp;
            <input type="radio" name="vCust_InfoEditProfile" value="1" <%=fcheck(fsqlboolean(vcust_infoeditprofile), 1)%> checked>Yes
            <input type="radio" name="vCust_InfoEditProfile" value="0" <%=fcheck(fsqlboolean(vcust_infoeditprofile), 0)%>>No&nbsp;&nbsp;&nbsp; Can Learner Edit their Profile?<br>&nbsp;&nbsp;&nbsp;&nbsp; Generally set to NO for Corporate.&nbsp;
            <br>
            <br>&nbsp;&nbsp;&nbsp;
            <input type="checkbox" name="vCust_VuNews" value="1" <%=fcheck(fsqlboolean(vcust_vunews), 1)%>>VuNews&nbsp;&nbsp;&nbsp;&nbsp; (Defaults to OFF and will only appear for EN users)<br>&nbsp;&nbsp;&nbsp;
            <input type="checkbox" name="vCust_Scheduler" value="1" <%=fcheck(fsqlboolean(vcust_scheduler), 1)%>>Scheduler&nbsp; (Defaults to OFF. This can be turned on via Tab1 and/or Tab 4)
          </td>
        </tr>
        <tr>
          <th>&nbsp;Tab 2 :</th>
          <td>
            <input type="checkbox" name="vCust_Tab2" value="1" <%=fcheck(fsqlboolean(vcust_tab2), 1)%>>My Learning
            <br>&nbsp;&nbsp;&nbsp;&nbsp; Name :
            <input type="text" name="vCust_Tab2Name" size="60" value="<%=vCust_Tab2Name%>" maxlength="62">
            <table class="table">
              <tr>
                <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
                <td>
                  <p class="c6">
                    <input type="text" name="vCust_MyWorldLaunch" size="20" value="<%=vCust_MyWorldLaunch%>"><br>Launch with this page. Note: precede with // if external link, else page must be in Repository/Tools, ie default.asp.&nbsp; Ensure learner has a way to return to My Learning - typically via the &quot;Globe&quot; icon/link.
                  </p>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <th>&nbsp;Tab 3 :</th>
          <td>
            <input type="checkbox" name="vCust_Tab3" value="1" <%=fcheck(fsqlboolean(vcust_tab3), 1)%>>My Content
            <br>&nbsp;&nbsp;&nbsp;&nbsp; Name :
            <input type="text" name="vCust_Tab3Name" size="60" value="<%=vCust_Tab3Name%>" maxlength="62">
          </td>
        </tr>
        <tr>
          <th>&nbsp;Tab 4 :</th>
          <td>
            <input type="checkbox" name="vCust_Tab4" value="1" <%=fcheck(fsqlboolean(vcust_tab4), 1)%>>Custom Tab<br>&nbsp;&nbsp;&nbsp;&nbsp; Name :
            <input type="text" name="vCust_Tab4Name" size="60" value="<%=vCust_Tab4Name%>" maxlength="62">
            <table class="table">
              <tr>
                <td nowrap width="50" rowspan="4">&nbsp;</td>
                <td nowrap colspan="3">
                  <br>Click to use Tab 4 then select one of the following functions :<br>
                  <input type="radio" name="vCust_Tab4Type" value="CL" <%=fcheck(vcust_tab4type, "cl")%>>Classroom Manager<br>
                  <input type="radio" name="vCust_Tab4Type" value="DF" <%=fcheck(vcust_tab4type, "df")%>>Discussion Forum<br>
                  <input type="radio" name="vCust_Tab4Type" value="SC" <%=fcheck(vcust_tab4type, "sc")%>>Scheduler<br>
                  <input type="radio" name="vCust_Tab4Type" value="RC" <%=fcheck(vcust_tab4type, "rc")%>>Resource Centre (include fields below) :<br></td>
              </tr>
              <tr>
                <td nowrap width="30" rowspan="3">&nbsp;</td>
                <td nowrap style="text-align: right">Programs : </td>
                <td>
                  <input type="text" name="vCust_Resources" size="40" value="<%=vCust_Resources%>"><br>Ie P1234EN P1235EN
                </td>
              </tr>
              <tr>
                <td nowrap style="text-align: right">Max Facilitators :</td>
                <td>
                  <input type="text" name="vCust_ResourcesMaxSponsor1" size="4" value="<%=vCust_ResourcesMaxSponsor%>"><br>This is the maximum number of facilitators that this account manager can issue to operate the Resource Centre (default 100).
                </td>
              </tr>
              <tr>
                <td nowrap style="text-align: right">Max Learners :</td>
                <td>
                  <input type="text" name="a" size="4"><br>This is the total number of learners that each facilitator can invite into the Resource Centre (default 100).
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <th>&nbsp;Tab 5 :</th>
          <td>
            <input type="checkbox" name="vCust_Tab5" value="1" <%=fcheck(fsqlboolean(vcust_tab5), 1)%>>More Content
            <br>&nbsp;&nbsp;&nbsp;&nbsp; Name :
            <input type="text" name="vCust_Tab5Name" size="60" value="<%=vCust_Tab5Name%>" maxlength="62">
          </td>
        </tr>
        <tr>
          <th>&nbsp;Tab 6 :</th>
          <td>
            <input type="checkbox" name="vCust_Tab6" value="1" <%=fcheck(fsqlboolean(vcust_tab6), 1)%>>Administration
            <br>&nbsp;&nbsp;&nbsp;&nbsp; Name :
            <input type="text" name="vCust_Tab6Name" size="60" value="<%=vCust_Tab6Name%>" maxlength="62">
          </td>
        </tr>
        <tr>
          <th>&nbsp;Tab 7 :</th>
          <td>
            <input type="checkbox" name="vCust_Tab7" value="1" <%=fcheck(fsqlboolean(vcust_tab7), 1)%>>Sign Off (normally leave on)
            <br>&nbsp;&nbsp;&nbsp;&nbsp; Name :
            <input type="text" name="vCust_Tab7Name" size="60" value="<%=vCust_Tab7Name%>" maxlength="62">
          </td>
        </tr>
      </table>
    </div>

    <div>
      <p style="text-align: right">
        <br>
        <a href="#Top">
          <img border="0" src="../Images/Icons/Ascending.gif" width="9" height="7"></a></p>
      <a <%=fstatx%> name="Bottom"></a>

      <div style="text-align: center;">

        <% If svMembLevel = 5 And svCustId <> vCust_Id Then %>

        <% If fHasLinkedCust(vCust_AcctId) And vCust_AcctId = Right(vCust_Id, 4) Then %>
        <div class="red" style="margin-bottom: 20px;">IMPORTANT: You cannot <b>Delete</b> this Customer profile because there are Linked Accounts using it. You must delete any Linked Accounts first.</div>
        <% ElseIf fHasLinkedCust(vCust_AcctId) And vCust_AcctId <> Right(vCust_Id, 4) Then %>
        <div class="red" style="margin-bottom: 20px;">IMPORTANT: This is a Linked Account. Deleting it will NOT delete supporting records (logs, etc).</div>
        <% ElseIf fHasChildCust(vCust_AcctId) Then %>
        <div class="red" style="margin-bottom: 20px;">IMPORTANT: You cannot <b>Delete</b> this Customer profile because it is a Parent Account with Children Accounts using it. You must delete any Child Accounts first.</div>
        <% Else %>
        <div class="red" style="margin-bottom: 20px;">IMPORTANT: If you <b>Delete</b> this Customer profile you will also delete ALL supporting records (logs, etc).</div>
        <% End If %>
        <% End If %>

        <div class="red" style="margin-bottom: 20px;">Note: <b>Return</b> does NOT Update this Profile.</div>

        <input type="submit" value="Update" name="bUpdate" class="button070"><%=f10%>

        <% If svMembLevel = 5 And svCustId <> vCust_Id Then %>

        <%  '...this is a linked child 
              If fHasLinkedCust(vCust_AcctId) And vCust_AcctId <> Right(vCust_Id, 4) Then %>
        <input onclick="javascript: jconfirm('Customer.asp?vDelCustId=<%=vCust_Id%>&amp;vFunction=del', 'Ok to delete this Customer but not the supporting files?\nThis action is irreversible.')" type="button" value="Delete" name="bDelete" class="button070"><%=f10%>
        <% '...this is for parents without children (linke or otherwise)
              ElseIf Not fHasChildCust(vCust_AcctId) AND Not fHasLinkedCust(vCust_AcctId) Then %>
        <input onclick="javascript: jconfirm('Customer.asp?vDelCustId=<%=vCust_Id%>&amp;vDelCustAcctId=<%=vCust_AcctId%>&amp;vFunction=del', 'Ok to delete this Account and all supporting files?\nThis action is irreversible.')" type="button" value="Delete" name="bDelete" class="button070"><%=f10%>
        <% End If %>

        <% End If %>

        <% Dim vUrl : vUrl = fIf (Len(vNext) > 0, vNext, "Customers.asp?vCustId=" & vCust_Id) %>

        <input onclick="location.href = '<%=vUrl%>'" type="button" value="Return" name="bAdd" class="button070">
      </div>

    </div>

  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
