<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/CustomCertRoutines.asp"-->
<!--#include virtual = "V5/Inc/Certificate.asp"-->

<% 
  Dim vCustId, vAddProgId, vEditProgId, vFunction, vMods, vRange, vLingo
  Dim aMods, Module

  vFunction = ""
  
  vLingo = fDefault(Request("vLingo"), "EN, FR, ES, PT")
  vRange = fDefault(Request("vRange"), "")

  '...update tables
  If Request("vFunction") = "add" Then
    sExtractProg
    vMods = Trim(fModsOk(vProg_Mods)) 
    If Len(vMods) > 0 Then 
      Response.Redirect "Error.asp?vErr=One or more module(s) are not on file (" & vMods & ").  Cannot Update."
    End If
    sInsertProg
    If Not vProg_Ok Then
      Response.Redirect "Error.asp?vErr=That Program is already on file !"
    End If
    Response.Redirect "Programs.asp?vRange=" & vRange & "&vLingo=" & vLingo

  ElseIf Request.Form("vFunction") = "edit" Then
    sExtractProg
    vMods = Trim(fModsOk(vProg_Mods)) 
    If Len(vMods) > 0 Then 
      Response.Redirect "Error.asp?vErr=One or more module(s) are not on file (" & vMods & ").  Cannot Update."
    End If
    sUpdateProg
    Response.Redirect "Programs.asp?vRange=" & vRange & "&vLingo=" & vLingo

  ElseIf Request.Form("vFunction") = "clone" Then
    If Not fCloneProg (Request("vProgId"), Request("vCloneId")) Then
      Response.Redirect "Error.asp?vErr=Either the source or object Program Id was invalid.  Cannot Clone."
    End If
    vProg_Id = Request("vProgId")

  ElseIf Request.Form("vFunction") = "cert" Then
    Response.Redirect fCertificateUrl("", "", 80, "", "", "Sample Module Title", "", fDefault(Request("vCust"), svCustId), fDefault(Request("vAcct"), svCustAcctId), Request("vProgId"), "", "", "")

  ElseIf Len(Request("vDelProgId")) = 7 Then 
    vProg_Id = Request("vDelProgId")
    sDeleteProg
    Response.Redirect "Programs.asp?vRange=" & vRange & "&vLingo=" & vLingo

  ElseIf Len(Request.Form("vAddProgID")) = 7 Then 
    vProg_Id = Request.Form("vAddProgID")
    vFunction = "add"

  ElseIf Len(Request.QueryString("vEditProgID")) = 7 Then 
    vProg_Id = Request.QueryString("vEditProgID")
    vFunction = "edit"

  End If

  '...get the values (even if trying to add)
  sGetProg (vProg_Id)

  '...if cloning then reset the prog Id
  If Request.Form("vFunction") = "clone" Then
    vProg_Id = Request("vCloneId")
    vRange = vProg_Id
    vFunction = "edit"
  End If
      
  '...set defaults
  vProg_Bookmark         = fDefault(vProg_Bookmark, "Y")
  vProg_CompletedButton  = fDefault(vProg_CompletedButton, "N")    
  vProg_LogTestResults   = fDefault(vProg_LogTestResults, "N")      
      
%>

<html>

<head>
  <title>Program</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>

  <script src="/V5/ckEditor/ckEditor.js"></script>
  <script src="/V5/ckEditor/ckEditorVu.js"></script>
  <script>$(function() { initCkEditorVu(); });</script>

  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>
  function validate(theForm) {

    if (theForm.vProg_Title2.value.length > 8000) {
      alert("Please enter at most 8000 characters in the \"Title 2\" field.");
      theForm.vProg_Title2.focus();
      return (false);
    }

    if (theForm.vProg_Owner.value == "") {
      alert("Please enter a value for the \"Owner Id\" field.");
      theForm.vProg_Owner.focus();
      return (false);
    }

    if (theForm.vProg_Owner.value.length < 4) {
      alert("Please enter at least 4 characters in the \"Owner Id\" field.");
      theForm.vProg_Owner.focus();
      return (false);
    }

    if (theForm.vProg_Owner.value.length > 4) {
      alert("Please enter at most 4 characters in the \"Owner Id\" field.");
      theForm.vProg_Owner.focus();
      return (false);
    }

    var checkOK = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyzƒŠŒŽšœžŸÀÁÂÃÄÅÆÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖØÙÚÛÜÝÞßàáâãäåæçèéêëìíîïðñòóôõöøùúûüýþÿ";
    var checkStr = theForm.vProg_Owner.value;
    var allValid = true;
    var validGroups = true;
    for (i = 0; i < checkStr.length; i++) {
      ch = checkStr.charAt(i);
      for (j = 0; j < checkOK.length; j++)
        if (ch == checkOK.charAt(j))
          break;
      if (j == checkOK.length) {
        allValid = false;
        break;
      }
    }
    if (!allValid) {
      alert("Please enter only letter characters in the \"Owner Id\" field.");
      theForm.vProg_Owner.focus();
      return (false);
    }

    if (theForm.vProg_EcomSplitOwner1.value.length > 5) {
      alert("Please enter at most 5 characters in the \"Ecommerce Owner Split1\" field.");
      theForm.vProg_EcomSplitOwner1.focus();
      return (false);
    }

    var checkOK = "0123456789-.";
    var checkStr = theForm.vProg_EcomSplitOwner1.value;
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
      alert("Please enter only digit characters in the \"Ecommerce Owner Split1\" field.");
      theForm.vProg_EcomSplitOwner1.focus();
      return (false);
    }

    if (decPoints > 1 || !validGroups) {
      alert("Please enter a valid number in the \"vProg_EcomSplitOwner1\" field.");
      theForm.vProg_EcomSplitOwner1.focus();
      return (false);
    }

    var chkVal = allNum;
    var prsVal = parseFloat(allNum);
    if (chkVal != "" && !(prsVal <= 100 && prsVal >= 0)) {
      alert("Please enter a value less than or equal to \"100\" and greater than or equal to \"0\" in the \"Ecommerce Owner Split1\" field.");
      theForm.vProg_EcomSplitOwner1.focus();
      return (false);
    }

    if (theForm.vProg_EcomSplitOwner2.value.length > 5) {
      alert("Please enter at most 5 characters in the \"Ecommerce Owner Split2\" field.");
      theForm.vProg_EcomSplitOwner2.focus();
      return (false);
    }

    var checkOK = "0123456789-.";
    var checkStr = theForm.vProg_EcomSplitOwner2.value;
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
      alert("Please enter only digit characters in the \"Ecommerce Owner Split2\" field.");
      theForm.vProg_EcomSplitOwner2.focus();
      return (false);
    }

    if (decPoints > 1 || !validGroups) {
      alert("Please enter a valid number in the \"vProg_EcomSplitOwner2\" field.");
      theForm.vProg_EcomSplitOwner2.focus();
      return (false);
    }

    var chkVal = allNum;
    var prsVal = parseFloat(allNum);
    if (chkVal != "" && !(prsVal <= 100 && prsVal >= 0)) {
      alert("Please enter a value less than or equal to \"100\" and greater than or equal to \"0\" in the \"Ecommerce Owner Split2\" field.");
      theForm.vProg_EcomSplitOwner2.focus();
      return (false);
    }

    var checkOK = "0123456789-.";
    var checkStr = theForm.vProg_US_Memo.value;
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
      alert("Please enter only digit characters in the \"US Prices\" field.");
      theForm.vProg_US_Memo.focus();
      return (false);
    }

    if (decPoints > 1 || !validGroups) {
      alert("Please enter a valid number in the \"vProg_US_Memo\" field.");
      theForm.vProg_US_Memo.focus();
      return (false);
    }

    var chkVal = allNum;
    var prsVal = parseFloat(allNum);
    if (chkVal != "" && !(prsVal <= 999 && prsVal >= 0)) {
      alert("Please enter a value less than or equal to \"999\" and greater than or equal to \"0\" in the \"US Prices\" field.");
      theForm.vProg_US_Memo.focus();
      return (false);
    }

    var checkOK = "0123456789-.";
    var checkStr = theForm.vProg_CA_Memo.value;
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
      alert("Please enter only digit characters in the \"CA Prices\" field.");
      theForm.vProg_CA_Memo.focus();
      return (false);
    }

    if (decPoints > 1 || !validGroups) {
      alert("Please enter a valid number in the \"vProg_CA_Memo\" field.");
      theForm.vProg_CA_Memo.focus();
      return (false);
    }

    var chkVal = allNum;
    var prsVal = parseFloat(allNum);
    if (chkVal != "" && !(prsVal <= 999 && prsVal >= 0)) {
      alert("Please enter a value less than or equal to \"999\" and greater than or equal to \"0\" in the \"CA Prices\" field.");
      theForm.vProg_CA_Memo.focus();
      return (false);
    }

    var checkOK = "0123456789-";
    var checkStr = theForm.vProg_Duration_Memo.value;
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
      alert("Please enter only digit characters in the \"Duration (Days)\" field.");
      theForm.vProg_Duration_Memo.focus();
      return (false);
    }

    var chkVal = allNum;
    var prsVal = parseInt(allNum);
    if (chkVal != "" && !(prsVal <= 999 && prsVal >= 0)) {
      alert("Please enter a value less than or equal to \"999\" and greater than or equal to \"0\" in the \"Duration (Days)\" field.");
      theForm.vProg_Duration_Memo.focus();
      return (false);
    }

    var radioSelected = false;
    for (i = 0; i < theForm.vProg_Bookmark.length; i++) {
      if (theForm.vProg_Bookmark[i].checked)
        radioSelected = true;
    }
    if (!radioSelected) {
      alert("Please select one of the \"Bookmarks?\" options.");
      return (false);
    }

    var radioSelected = false;
    for (i = 0; i < theForm.vProg_CompletedButton.length; i++) {
      if (theForm.vProg_CompletedButton[i].checked)
        radioSelected = true;
    }
    if (!radioSelected) {
      alert("Please select one of the \"Completed Button?\" options.");
      return (false);
    }

    var radioSelected = false;
    for (i = 0; i < theForm.vProg_TaxExempt.length; i++) {
      if (theForm.vProg_TaxExempt[i].checked)
        radioSelected = true;
    }
    if (!radioSelected) {
      alert("Please select one of the \"Tax Exempt?\" options.");
      return (false);
    }

    if (theForm.vProg_Memo.value.length > 512) {
      alert("Please enter at most 512 characters in the \"vProg_Memo\" field.");
      theForm.vProg_Memo.focus();
      return (false);
    };

    if (theForm.vProg_Assessment.value.length > 0 && theForm.vProg_Assessment.value.length < 6 ) {
      alert("Please enter a valid Launch Module Id.");
      theForm.vProg_Assessment.focus();
      return (false);
    };


    if (theForm.vProg_AssessmentAttempts.value == "") {
      alert("Please enter a value for the \"Max Attempts\" field.");
      theForm.vProg_AssessmentAttempts.focus();
      return (false);
    }

    if (theForm.vProg_AssessmentAttempts.value.length < 1) {
      alert("Please enter at least 1 characters in the \"Max Attempts\" field.");
      theForm.vProg_AssessmentAttempts.focus();
      return (false);
    }

    if (theForm.vProg_AssessmentAttempts.value.length > 2) {
      alert("Please enter at most 2 characters in the \"Max Attempts\" field.");
      theForm.vProg_AssessmentAttempts.focus();
      return (false);
    }

    var checkOK = "0123456789-";
    var checkStr = theForm.vProg_AssessmentAttempts.value;
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
      theForm.vProg_AssessmentAttempts.focus();
      return (false);
    }

    var chkVal = allNum;
    var prsVal = parseInt(allNum);
    if (chkVal != "" && !(prsVal <= 99 && prsVal >= 0)) {
      alert("Please enter a value less than or equal to \"99\" and greater than or equal to \"0\" in the \"Max Attempts\" field.");
      theForm.vProg_AssessmentAttempts.focus();
      return (false);
    }

    var checkOK = "0123456789-.";
    var checkStr = theForm.vProg_AssessmentScore.value;
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
      theForm.vProg_AssessmentScore.focus();
      return (false);
    }

    if (decPoints > 1 || !validGroups) {
      alert("Please enter a valid number in the \"vProg_AssessmentScore\" field.");
      theForm.vProg_AssessmentScore.focus();
      return (false);
    }

    var chkVal = allNum;
    var prsVal = parseFloat(allNum);
    if (chkVal != "" && !(prsVal <= 1 && prsVal >= 0)) {
      alert("Please enter a value less than or equal to \"1\" and greater than or equal to \"0\" in the \"Assessment Score\" field.");
      theForm.vProg_AssessmentScore.focus();
      return (false);
    }
    return (true);
  }





  </script>
</head>

<body>
  <% 
    Server.Execute vShellHi
  %>

  <h1>Program Details</h1>
  <h2>Modify any Program values then click <b>Update</b>.</h2>

  <form id="fProgram" method="POST" action="Program.asp" target="_self" onsubmit="return validate(this)">

    <input type="hidden" name="vFunction" value="<%=vFunction%>">
    <input type="hidden" name="vProg_Id" value="<%=vProg_Id%>">
    <input type="hidden" name="vProg_Length" value="<%=vProg_Length%>">
    <input type="hidden" name="vRange" value="<%=vRange%>">
    <input type="hidden" name="vLingo" value="<%=vLingo%>">

    <table class="table">
      <tr>
        <th>Program Id :</th>
        <td><%=vProg_Id%> <% = fIf (svMembLevel = 5, f10 & "[Internal Program No : " & vProg_No & "]", "") %> </td>
      </tr>
      <tr>
        <th>Title1 :</th>
        <td>
          <input type="text" size="71" name="vProg_Title1" value="<%=vProg_Title1%>" maxlength="256">
        </td>
      </tr>
      <tr>
        <th>Title2 :</th>
        <td>
          <textarea name="vProg_Title2"><%=vProg_Title2%></textarea>
          <br>
          For alternate or longer titles, enter 8 char Customer Id, space then the variation separated from the next title by a tilde, ie: CMAC2250 This is a new title~CFIB0001 This is yet another title~CFIB0002 Third Title Variation
        </td>
      </tr>
      <tr>
        <th>Promo :</th>
        <td>
          <textarea name="vProg_Promo"><%=vProg_Promo%></textarea>
          <br>
          Enter any promotional text to appear in More Content, italicized in red below the title as follows:
          <br>
          <span class="red"><i>Do not enter any HTML tags.</i></span>
        </td>
      </tr>
      <tr>
        <th>Retired :</th>
        <td>
          <input type="radio" value="0" name="vProg_Retired" <%=fcheck("0", fsqlboolean(vprog_retired))%>>No
          <input type="radio" value="1" name="vProg_Retired" <%=fcheck("1", fsqlboolean(vprog_retired))%>>Yes
          <br>
          A retired program is still accessible and active but cannot be sold via Ecommerce even if it is on the clients catalogue. This is similar to changing the price of a Program to 9999. 
        </td>
      </tr>
      <tr>
        <th>Owner Id :</th>
        <td>
          <input type="text" name="vProg_Owner" size="14" value="<%=fDefault(vProg_Owner, "VUBZ")%>" maxlength="4">
          <br>
          Enter VUBZ for Vubiz content, else 4 character channel group code, ie: TELE (no spaces or quotes).&nbsp; This field allows partners to view their content sales in the ecommerce report. Ensure you use the same Owner Id for each program owned by this owner. The code should be the first 4 character of the owner&#39;s account Id. The Ecom Owner summary will show owner revenue against this 4 character code.
          <div style="text-align: right">
            <table class="table">
              <tr>
                <td>
                  The following ecommerce splits apply to owner accounts starting with first four characters of Owner Id, ie &quot;TELE&quot; would embrace TELE1234 and TELE1254.<br />
                  &nbsp;<input type="text" name="vProg_EcomSplitOwner1" size="4" value="<%=vProg_EcomSplitOwner1%>" maxlength="5"> % Owner Split if sold by Owner Account(s)<br>
                  &nbsp;<input type="text" name="vProg_EcomSplitOwner2" size="4" value="<%=vProg_EcomSplitOwner2%>" maxlength="5"> % Owner Split if sold by Other Accounts
                </td>
              </tr>
            </table>
          </div>
        </td>
      </tr>
      <tr>
        <th>Description :</th>
        <td>
          <textarea id="editor" maxlength="8000" name="vProg_Desc"><%=vProg_Desc%></textarea>
          <script>
            CKEDITOR.replace('vProg_Desc');
            CKEDITOR.add
            CKEDITOR.config.contentsCss = '/V5/Inc/Vubi2.css';  
          </script>
        </td>
      </tr>
      <tr>
        <th>Module Ids :</th>
        <td>
          <textarea name="vProg_Mods"><%=vProg_Mods%></textarea>
          <% 
            If Len(vProg_Mods) > 0 Then 
          %>
          <br>
          Enter Module Ids separated by spaces, ie: 0023EN 0101EN.&nbsp; Click to access modules:
          <br>
          <% 
              aMods = Split(vProg_Mods, " ")
              For i = 0 to Ubound(aMods)
                vMods_Id = aMods(i)
                Select Case fModsStatus (vMods_Id)
                  Case 0 : Module = "<font color='black'>" & vMods_Id & "</font>"
                  Case 1 : Module = vMods_Id
                  Case 2 : Module = "<font color='red'>" & vMods_Id & "</font>"
                  Case 3 : Module = "<font color='orange'>" & vMods_Id & "</font>"
                End Select
          %> <a target="_blank" href="Module.asp?vMods_Id=<%=vMods_Id%>"><%=Module%></a>
          <%
  	          Next
            End If 
          %>
          <br>
          [Black: Not on file, Red: Inactive, Blue: Active/Ok for Completion, Yellow: Not for Completion]
        </td>
      </tr>
      <tr>
        <th>Multi-SCO Children&nbsp;&nbsp;
          <br>
          Modules Ids : </th>
        <td>
          <textarea name="vProg_Scos"><%=vProg_Scos%></textarea><br>
          If this is an FX Multi-SCO program and a single Parent Mod Id appears in the Module Ids field above, then enter the Children Module Ids separated by spaces, ie: 0023EN 0101EN. 
          <% 
            If Len(vProg_Scos) > 0 Then 
          %>  Click to access modules:
          <br>
          <% 
              aMods = Split(vProg_Scos, " ")
              For i = 0 to Ubound(aMods)
                vMods_Id = aMods(i)
          %> <a target="_blank" href="Module.asp?vMods_Id=<%=vMods_Id%>"><%=vMods_Id%></a>
          <%
  	          Next
            End If 
          %>
        </td>
      </tr>
      <tr>
        <th>Suggested US Price :</th>
        <td>
          <input type="text" name="vProg_US_Memo" size="9" value="<%=vProg_US_Memo%>">
          US Dollars (this can be used for CDs or can override the customer ecom string)
        </td>
      </tr>
      <tr>
        <th>Suggested CA Price :</th>
        <td>
          <input type="text" name="vProg_CA_Memo" size="9" value="<%=vProg_CA_Memo%>">
          CA Dollars (this can be used for CDs or can override the customer ecom string)
        </td>
      </tr>
      <tr>
        <th>Suggested Duration :</th>
        <td>
          <input type="text" name="vProg_Duration_Memo" size="9" value="<%=vProg_Duration_Memo%>">
          Days - use 999 for CDs
        </td>
      </tr>
      <tr>
        <th>Group 1 Pricing :<br>
          &nbsp;</th>
        <td>When sold via ecommerce, enter group 1 pricing.&nbsp; This overrides the values set in the Customer profile.<br>
          <input type="text" name="vProg_EcomGroupLicense" size="6" value="<%=vProg_EcomGroupLicense%>">
          Annual License as Ratio of Individual Pricing. Note: 3.0 if left empty.&nbsp; To force to 0 enter 0.0001.<br>
          <input type="text" name="vProg_EcomGroupSeat" size="6" value="<%=vProg_EcomGroupSeat%>">
          Per Seat Fee as Ratio of Individual Pricing.&nbsp; Note: 0.2 if left empty.&nbsp; To force to 0 enter 0.0001.
        </td>
      </tr>
      <tr>
        <th>Group 2 Discounts ? </th>
        <td>
          <input type="radio" name="vProg_Discounts" value="Y" <%=fcheck("y", vprog_discounts)%>>Yes (allow normal volume discounts - default)<br>
          <input type="radio" name="vProg_Discounts" value="N" <%=fcheck("n", vprog_discounts)%>>No&nbsp; (do NOT allow volume discounts for this program)
        </td>
      </tr>
      <tr>
        <th>Bookmark Modules ?</th>
        <td>
          <input type="radio" value="Y" name="vProg_Bookmark" <%=fcheck(vprog_bookmark, "y")%>>Yes
          <input type="radio" value="N" name="vProg_Bookmark" <%=fcheck(vprog_bookmark, "n")%>>No&nbsp; (disable for group/anonymous usage )
        </td>
      </tr>
      <tr>
        <th>Add Completed Button in fModules ?</th>
        <td>
          <input type="radio" value="N" name="vProg_CompletedButton" <%=fcheck(vprog_completedbutton, "n")%>>No&nbsp; (disable for ecommerce)<br>
          <input type="radio" value="Y" name="vProg_CompletedButton" <%=fcheck(vprog_completedbutton, "y")%>>Yes (enable for corporate)
        </td>
      </tr>
      <tr>
        <th>GST/HST Tax Exempt ?</th>
        <td>
          <input type="radio" name="vProg_TaxExempt" value="0" <%=fcheck(fsqlboolean(vprog_taxexempt), 0)%>>No
          <input type="radio" name="vProg_TaxExempt" value="1" <%=fcheck(fsqlboolean(vprog_taxexempt), 1)%>>Yes (do not charge GST and/or HST on Canadian sales)
        </td>
      </tr>

      <tr>
        <th>Length :</th>
        <td><%=vProg_Length%> (Hours - computed)</td>
      </tr>

      <tr>
        <th>NASBE CPE :</th>
        <td><input type="text" name="vProg_Nasba_Cpe" size="9" value="<%=vProg_Nasba_Cpe%>"> For certificates. Leave empty if not used (ie NOT 0).</td>         
      </tr>


      <tr>
        <th>Memo :</th>
        <td>
          <textarea rows="3" name="vProg_Memo"><%=vProg_Memo%></textarea>
        </td>
      </tr>
      <tr>
        <th>Recurring Assessment ?</th>
        <td>
          <input type="text" name="vProg_ResetStatus" size="3" value="<%=fDefault(vProg_ResetStatus, 0)%>">
          Days after which time the assessment must be retaken (ie 90, 180, 365). Leave at 0 if the status does not get refreshed. If, for example, an assessment must be taken every 180 days then any completed assessment in this Program will revert to incomplete after 180 days. <b><font color="#FF0000">Caution: changing this value on existing programs can render History records as suspect.</font></b>
        </td>
      </tr>
      <tr>
        <th colspan="2">
          <h2>
            <br>
            Assessment details ...</h2>
        </th>
      </tr>
      <tr>
        <th>Launch Module :</th>
        <td>
          <input type="text" size="8" name="vProg_Assessment" value="<%=vProg_Assessment%>" maxlength="9">
          <br>
          Leave empty or enter the Module Id that will launch this assessment.<br>
          <span class="red">If this is the only Module in this Program it must be entered in the 'Module Ids' text box above.</span>
        </td>
      </tr>
      <tr>
        <th>Max Attempts :</th>
        <td>
          <input type="text" size="4" name="vProg_AssessmentAttempts" value="<%=fDefault(vProg_AssessmentAttempts, 0)%>" maxlength="2">
          If the max attempts for any VuAssess Test is other than the default of 3 (in which case leave as 0), specify the maximum no of attempts&nbsp; where 99 specifies no restrictions.
        </td>
      </tr>
      <tr>
        <th>Passing Score :</th>
        <td>
          <input type="text" name="vProg_AssessmentScore" size="4" value="<%=fDefault(vProg_AssessmentScore, 0)%>">
          Score needed to display a certificate (must be between 0 and 1 - ie 70% is entered as .7).&nbsp; If you enter zero then 80% is assumed (.8).&nbsp; If you wish to display a certificate regardless of score then enter .01.
        </td>
      </tr>
      <tr>
        <th>Certificate (Advanced) :</th>
        <td>
          <textarea rows="3" name="vProg_AssessmentIds"><%=vProg_AssessmentIds%></textarea><br>
          Generate a Certificate of Completion if all of the above Assessments (Program 
          Id | Module Id) are flagged as Completed on the RTE. Enter as: &quot;P1234EN|1234EN P1234EN|1235EN P9876EN|0022EN&quot;. The Modules can come from any number of Programs. Typically this Program&#39;s sole purpose is to represent the certificate.
        </td>
      </tr>
      <tr>
        <th>Certificate (Basic) :</th>
        <td>
          <input type="radio" value="0" name="vProg_CertSimple" <%=fcheck(fsqlboolean(vprog_certsimple), 0)%>>No
          <input type="radio" value="1" name="vProg_CertSimple" <%=fcheck(fsqlboolean(vprog_certsimple), 1)%>>Yes<br>
          Generate a Certificate of Completion if the RTE deems the Program is complete.<br />
					Note: this feature is only functional on V8 - not V5.
        </td>
      </tr>
      <tr>
        <td colspan="2" style="text-align: center; padding-top: 30px;">

          <!--         <input type="submit" value="Delete" name="bDelete" class="button" onclick="jconfirm('Program.asp?vDelProgID=<%=vProg_Id%>&vFunction=del', 'Ok to delete?')">-->
          <input type="button" value="Delete" name="bDelete" class="button" onclick="jconfirm('Program.asp?vDelProgID=<%=vProg_Id%>&vFunction=del', 'Ok to delete?')">


          <%=f10%>
          <input type="submit" value="Update" name="bUpdate" class="button">
          <br>
          <br>
          <a href="Programs.asp?vRange=<%=vRange%>&vLingo=<%=vLingo%>">Program List</a>
        </td>
      </tr>
    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
