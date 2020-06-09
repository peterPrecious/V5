<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% 
  Dim vCertLang, vTitle, vFirstName, vLastName, vScript, vScore, vLastScore, vUrl, vFolder, logo

  vScript     = ""

  vCertLang   = fDefault(Request("vCertLang"), "EN")
  vFirstName  = Request("vFirstName")
  vLastName   = Request("vLastName")

  vProg_Id    = Request("vProg_Id")  
  vTitle      = Request("vTitle")   
  vMods_Id    = Request("vMods_Id")   

  vScore      = Request("vScore")   
  vLastScore  = fFormatDate(Request("vLastScore"))


  '...Postback? 
  If Request.Form("vHidden") = "y" Then    

    sGetProg (vProg_Id) '...can be invalid, get custom cert path

    sGetMods (vMods_Id) '...can be invalid
    If Len(vTitle) = 0 And Len(vMods_Title) > 0 Then
      vTitle = vMods_Title
    End If
    
    sGetCust (svCustId) '...will always be valid, get custom cert path

    If Len(vProg_AssessmentCert) > 0 Then
      vFolder = vProg_AssessmentCert        
    ElseIf Len(vCust_AssessmentCert) > 0 Then
      vFolder = vCust_AssessmentCert        
    Else
      vFolder = vCertLang
    End If

    logo = fIf(Len(svCustBanner) > 0, svCustBanner, "")

    vUrl = ""
    vUrl = vUrl & "/v5/Assessments/CustomCerts/" & vFolder & "/default.htm" 
    vUrl = vUrl & "?vMods_Id="    & vMods_Id
    vUrl = vUrl & fIf(Len(vTitle) > 0     , "&vMods_Title=" & fjUnquote(vTitle), "")
    vUrl = vUrl & fIf(Len(vFirstName) > 0 , "&vFirstName="  & fjUnquote(vFirstName), "")
    vUrl = vUrl & fIf(Len(vLastName) > 0  , "&vLastName="   & fjUnquote(vLastName), "")
   
    If Len(vScore) > 0 Then
      vScore = Cint(vScore)
      If vScore = 100 Then
        vUrl = vUrl & "&vScore=1"
      Else
        vUrl = vUrl & "&vScore=." & vScore
      End If
    End If

    vUrl = vUrl & fIf(Len(vLastScore) > 0  , "&vLastScore="   & vLastScore, "")
    
    vScript = ""
    vScript = vScript & "<script>" & vbCrLf
    vScript = vScript & "  window.open('" & vUrl & "','Certificate','width=750,height=475,left=100,top=100,status=no,scrollbars=no,resizable=no')" & vbCrLf
    vScript = vScript & "</script>"
    
  End If
%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <%=vScript%>
  <title>Certificate Generator</title>
</head>

<body>

  <% Server.Execute vShellHi %>  

  <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vFirstName.value == "")
  {
    alert("Please enter a value for the \"First Name\" field.");
    theForm.vFirstName.focus();
    return (false);
  }

  if (theForm.vFirstName.value.length > 32)
  {
    alert("Please enter at most 32 characters in the \"First Name\" field.");
    theForm.vFirstName.focus();
    return (false);
  }

  if (theForm.vLastName.value == "")
  {
    alert("Please enter a value for the \"Last Name\" field.");
    theForm.vLastName.focus();
    return (false);
  }

  if (theForm.vLastName.value.length > 32)
  {
    alert("Please enter at most 32 characters in the \"Last Name\" field.");
    theForm.vLastName.focus();
    return (false);
  }

  if (theForm.vMods_Id.value == "")
  {
    alert("Please enter a value for the \"Module ID\" field.");
    theForm.vMods_Id.focus();
    return (false);
  }

  if (theForm.vMods_Id.value.length < 6)
  {
    alert("Please enter at least 6 characters in the \"Module ID\" field.");
    theForm.vMods_Id.focus();
    return (false);
  }

  if (theForm.vMods_Id.value.length > 6)
  {
    alert("Please enter at most 6 characters in the \"Module ID\" field.");
    theForm.vMods_Id.focus();
    return (false);
  }

  if (theForm.vScore.value.length > 3)
  {
    alert("Please enter at most 3 characters in the \"Score\" field.");
    theForm.vScore.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.vScore.value;
  var allValid = true;
  var validGroups = true;
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
    alert("Please enter only digit characters in the \"Score\" field.");
    theForm.vScore.focus();
    return (false);
  }

  var chkVal = allNum;
  var prsVal = parseInt(allNum);
  if (chkVal != "" && !(prsVal >= 0 && prsVal <= 100))
  {
    alert("Please enter a value greater than or equal to \"0\" and less than or equal to \"100\" in the \"Score\" field.");
    theForm.vScore.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="CertificateGenerate.asp" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
    <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="2" cellspacing="0">
      <tr>
        <td valign="top" colspan="2"><h1 align="center">Generate a Certificate</h1><h2 align="center">This will create a certificates with the fields entered below using the logo for this account.<br>A custom certificate will be used if the account or program requests them.</h2></td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap width="50%">Language : </th>
        <td valign="top" width="50%">
          <input type="radio" name="vCertLang" value="EN" <%=fcheck(vCertLang, "EN")%>>EN&nbsp;&nbsp; 
          <input type="radio" name="vCertLang" value="FR" <%=fcheck(vCertLang, "FR")%>>FR&nbsp;&nbsp; 
          <input type="radio" name="vCertLang" value="ES" <%=fcheck(vCertLang, "ES")%>>ES</td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap width="50%">First Name : </th>
        <td valign="top" width="50%">&nbsp;<!--webbot bot="Validation" s-display-name="First Name" s-data-type="String" b-value-required="TRUE" i-maximum-length="32" --><input type="text" name="vFirstName" size="20" value="<%=vFirstName%>" maxlength="32"></td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap width="50%">Last Name : </th>
        <td valign="top" width="50%">&nbsp;<!--webbot bot="Validation" s-display-name="Last Name" s-data-type="String" b-value-required="TRUE" i-maximum-length="32" --><input type="text" name="vLastName" size="20" value="<%=vLastName%>" maxlength="32"></td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap width="50%">Course (Program Id) : </th>
        <td valign="top" width="50%">&nbsp;<input type="text" name="vProg_Id" size="13" value="<%=vProg_Id%>"> Optional<br>&nbsp;Only required if this program/module uses a custom certificate</td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap width="50%">Assessment (Module) ID : </th>
        <td valign="top" width="50%">&nbsp;<!--webbot bot="Validation" s-display-name="Module ID" b-value-required="TRUE" i-minimum-length="6" i-maximum-length="6" --><input type="text" name="vMods_Id" size="13" maxlength="6" value="<%=vMods_Id%>"> Mandatory</td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap width="50%">Assessment (Module) Name : </th>
        <td valign="top" width="50%">&nbsp;<input type="text" name="vTitle" size="30" value="<%=vTitle%>"><br>&nbsp;Leave empty if you use the Module ID above and the appropriate<br>&nbsp;Module Name/Title will appear</td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap width="50%">Score : </th>
        <td valign="top" width="50%">&nbsp;<!--webbot bot="Validation" s-display-name="Score" s-data-type="Integer" s-number-separators="x" i-maximum-length="3" s-validation-constraint="Greater than or equal to" s-validation-value="0" s-validation-constraint="Less than or equal to" s-validation-value="100" --><input type="text" name="vScore" size="5" maxlength="3" value="<%=vScore%>"> 0-100, leave empty if not used</td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap width="50%">Date of Assessment : </th>
        <td valign="top" width="50%">&nbsp;<input type="text" name="vLastScore" size="13" value="<%=fFormatDate(Now())%>"> ie Jan 15, 2008 (MMM, D, YYY)<br>Leave empty if not used</td>
      </tr>
      <tr>
        <td align="center" valign="top" colspan="2">&nbsp;<p><input type="submit" value="Generate" name="bGenerate" class="button"></p><p>&nbsp;</p></td>
      </tr>
    </table>
    <input type="hidden" name="vHidden" value="y">
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
