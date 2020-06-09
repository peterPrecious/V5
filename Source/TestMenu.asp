<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Test.asp"-->

<html>
<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

<% Server.Execute vShellHi %>

<table border="1" width="100%" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="0">
  <tr>
    <td width="100%" colspan="2">
    <h1>Edit Self Assessment Table</h1>
    <p>Either enter a new self assessment Id you wish to add (ie 1234EN) then click &quot;add&quot;, OR select an existing self assessment you wish to edit then click &quot;go&quot;<br>&nbsp;&nbsp; </p></td>
  </tr>
  <tr>
    <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vModId.value == "")
  {
    alert("Please enter a value for the \"Module/Self Assessment Id\" field.");
    theForm.vModId.focus();
    return (false);
  }

  if (theForm.vModId.value.length < 6)
  {
    alert("Please enter at least 6 characters in the \"Module/Self Assessment Id\" field.");
    theForm.vModId.focus();
    return (false);
  }

  if (theForm.vModId.value.length > 6)
  {
    alert("Please enter at most 6 characters in the \"Module/Self Assessment Id\" field.");
    theForm.vModId.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="TestEdit.asp" target="_self" name="FrontPage_Form1" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">
      <td width="50%" bgcolor="#DDEEF9" height="20">&nbsp;&nbsp;&nbsp;&nbsp; <!--webbot bot="Validation" s-display-name="Module/Self Assessment Id" b-value-required="TRUE" i-minimum-length="6" i-maximum-length="6" -->
        <input type="text" name="vModId" size="13" maxlength="6"> 
        <input border="0" src="../Images/Buttons/Add_<%=svLang%>.gif" name="I4" type="image"> 
      </td>
    </form>

    <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form2_Validator(theForm)
{

  if (theForm.vModId.selectedIndex < 0)
  {
    alert("Please select one of the \"Module Id\" options.");
    theForm.vModId.focus();
    return (false);
  }

  if (theForm.vModId.selectedIndex == 0)
  {
    alert("The first \"Module Id\" option is not a valid selection.  Please choose one of the other options.");
    theForm.vModId.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="TestEdit.asp" target="_self" onsubmit="return FrontPage_Form2_Validator(this)" name="FrontPage_Form2" language="JavaScript">
      <td width="50%" bgcolor="#DDEEF9" align="right" height="20">&nbsp;&nbsp;&nbsp;&nbsp; <!--webbot bot="ValIdation" s-display-name="Module Id" b-value-required="TRUE" b-disallow-first-item="TRUE" -->
        <select size="1" name="vModId">
        <option selected value="Select">Select Module</option>
        <%=fTestOptionsAll%>
        </select> 
        <input border="0" src="../Images/Buttons/Edit_<%=svLang%>.gif" name="I3" type="image"> 
      </td>
    </form>
  </tr>
</table>
<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body></html>
