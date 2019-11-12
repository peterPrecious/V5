<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include file = "Completion_Routines.asp"-->
<!--#include file = "Completion_LocationManager_Routines.asp"-->

<%
   '...Add a new Region/Location
  If Request("vReg").Count > 0 Then
    vUnit_L1        = Ucase(Trim(Request("vUnit_L1")))
    vUnit_L1Title   = Trim(Request("vUnit_L1Title"))
    vUnit_L0        = Ucase(Trim(Request("vUnit_L0")))
    vUnit_L0Title   = Trim(Request("vUnit_L0Title"))

    If fUnitExists (vUnit_L1, vUnit_L0) Then
      vMsg = "Region | " & Session("Completion_L0Tit") & " : " & vUnit_L1 & " | " & vUnit_L0 & "<br>are not unique!"
    Else      
      '...update the unit table
      sInsertUnit  vUnit_L1, vUnit_L1Title, vUnit_L0, vUnit_L0Title       
      '...add roles to crit table 
      sOpenDb2
      For i = 0 To Ubound(aRoleXX) 
        vSql = " INSERT INTO Crit (Crit_AcctId, Crit_Id) VALUES ('" & svCustAcctId & "', '" & vUnit_L1 & " " & vUnit_L0 & " " & aRoleXX(i) & "')"
        sCompletion_Debug
        oDb2.Execute(vSql)
      Next
      sCloseDb2
      vMsg = "Region | " & Session("Completion_L0Tit") & " : " & vUnit_L1 & " | " & vUnit_L0 & " was added successfully."
    End If

    sGetUnitByL0 vUnit_L0  '...get the new unit number
    Response.Redirect "Completion_LocationManager_Rev.asp?vUnit_No=" & vUnit_No
 
    
  End If
  
%>
<html>

<head>
  <title>Completion_LocationManager_Reg</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <script>
  
    // field tests
    var reAlphaNumeric = new RegExp(/^[0-9A-Za-z]+$/);
    var reAlpha        = new RegExp(/^[A-Za-z]+$/);
    var reNumeric      = new RegExp(/^[0-9]+$/);

    var L0len          = <%=Session("Completion_L0len")%>;
    var L1len          = <%=Session("Completion_L1len")%>;
    
    var L0tit          = "<%=Session("Completion_L0tit")%>";
    var L1tit          = "<%=Session("Completion_L1tit")%>";

    function validateReg (theForm) {
      if (theForm.vUnit_L1.value.length != L1len) {
        alert("Please enter a " + L1len + " character " + L1tit + " Id.");
        theForm.vUnit_L1.focus();
        return (false);
      }
      if (theForm.vUnit_L1Title.value == "") {
        alert("Please enter the " + L1tit + " Name.");
        theForm.vUnit_L1Title.focus();
        return (false);
      }    
      if (theForm.vUnit_L1Title.value.length < 4 || theForm.vUnit_L1Title.value.length > 128) {
        alert("The " + L1tit + " Name must be between 4 and 128 characters.");
        theForm.vUnit_L1Title.focus();
        return (false);  
      }
      if (theForm.vUnit_L0.value.length != L0len) {
        alert("Please enter a " + L0len + " character " + L0tit + " Id.");
        theForm.vUnit_L0.focus();
        return (false);
      }
      if (theForm.vUnit_L0Title.value == "") {
        alert("Please enter the " + L0tit + " Name.");
        theForm.vUnit_L0Title.focus();
        return (false);
      }    
      if (theForm.vUnit_L0Title.value.length < 4 || theForm.vUnit_L0Title.value.length > 128) {
        alert("The " + L0tit + " Name must be between 4 and 128 characters.");
        theForm.vUnit_L0Title.focus();
        return (false);  
      }
      return (true);
    }   
  </script>
</head>

<body>

  <% Server.Execute vShellHi %>
  <!--#include file = "Completion_LocationManager_Top.asp"-->

  <div style="margin-bottom: 30px;">
    <h1>Add an new Region</h1>
    <h2>When adding a new Region, you must include the initial <%=Session("Completion_L0Tit")%>.<br>Both the Region ID and the <%=Session("Completion_L0Tit")%> ID must be unique.</h2>
    <%=fIf(Len(vMsg)>0, "<h5>" & vMsg & "</h5>", "")%>
  </div>

  <form method="POST" action="Completion_LocationManager_Reg.asp" id="fReg" name="fReg" onsubmit="return validateReg(this)">
    <table style="width:600px; margin:auto">
      <tr>
        <td class="rowshade">Region ID</td>
        <td class="rowshade">Region Name</td>
        <td rowspan="3">&nbsp;</td>
      </tr>
      <tr>
        <td style="text-align:center"><input type="text" name="vUnit_L1" size="8"></td>
        <td style="text-align:center"><input type="text" name="vUnit_L1Title" size="40"></td>
      </tr>
      <tr>
        <td class="rowshade">&nbsp;<%=Session("Completion_L0Tit")%> ID</td>
        <td class="rowshade"><%=Session("Completion_L0Tit")%> Name</td>
      </tr>
      <tr>
        <td style="text-align:center"><input type="text" name="vUnit_L0" size="8"></td>
        <td style="text-align:center"><input type="text" name="vUnit_L0Title" size="40"></td>
        <td style="text-align:center"><input type="submit" value="Add" name="vReg" id="vReg" class="button"></td>
      </tr>
    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>
