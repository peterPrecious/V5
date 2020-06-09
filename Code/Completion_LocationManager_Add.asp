<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include file = "Completion_Routines.asp"-->
<!--#include file = "Completion_LocationManager_Routines.asp"-->

<%
  '...get the region values from the selected list before edditing
  '   <option value='0002|Distribution Centre|0'">0002 (Distribution Centre)</option>
  If Request.QueryString("vReg").Count > 0 Then
    i = Split(Request("vReg"), "|")
    vUnit_L1        = i(0)
    vUnit_L1Title   = i(1)
    vUnit_HO        = i(2)

 '...Add a new Region/Location
  ElseIf Request("vReg").Count > 0 Then
    vUnit_L1        = Ucase(Trim(Request("vUnit_L1")))
    vUnit_L1Title   = Trim(Request("vUnit_L1Title"))
    vUnit_L0        = Ucase(Trim(Request("vUnit_L0")))
    vUnit_L0Title   = Trim(Request("vUnit_L0Title"))

    If fLocnExists (vUnit_L0) Then
      vMsg = Session("Completion_L1tit") & " : " & vUnit_L0 & " already exists!"
    Else      
      '...update the unit table
      sInsertUnit  vUnit_L1, vUnit_L1Title, vUnit_L0, vUnit_L0Title       
      '...add roles to crit table 
      sOpenDb2

      If vUnit_HO = 0 Then
        For i = 0 To Ubound(aRoleXX) 
          vSql = " INSERT INTO Crit (Crit_AcctId, Crit_Id) VALUES ('" & svCustAcctId & "', '" & vUnit_L1 & " " & vUnit_L0 & " " & aRoleXX(i) & "')"
          sCompletion_Debug
          oDb2.Execute(vSql)
        Next
      Else
        For i = 0 To Ubound(aRoleHO) 
          vSql = " INSERT INTO Crit (Crit_AcctId, Crit_Id) VALUES ('" & svCustAcctId & "', '" & vUnit_L1 & " " & vUnit_L0 & " " & aRoleHO(i) & "')"
          sCompletion_Debug
          oDb2.Execute(vSql)
        Next
      End If
      sCloseDb2
      vMsg = Session("Completion_L0tit") & " : " & vUnit_L0 & " was added successfully."
    End If
  End If
  
%>

<html>

<head>
  <title>Completion_LocationManager_Add</title>
  <meta charset="UTF-8">
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

    function validateAdd (theForm) {
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
    <h1>Add an new <%=Session("Completion_L0Tit")%></h1>
    <h2>Specify a unique <%=Session("Completion_L0Tit")%> ID and Name then click Add</h2>
    <%=fIf(Len(vMsg)>0, "<h5>" & vMsg & "</h5>", "")%>
  </div>

  <form method="POST" action="Completion_LocationManager_Add.asp" id="fReg" name="fReg" onsubmit="return validateAdd(this)">

    <table style="width:600px; margin:auto;">
      <tr>
        <td class="rowshade">Region ID</td>
        <td class="rowshade">Region Name</td>
        <td rowspan="3">&nbsp;</td>
      </tr>
      <tr>
        <td style="text-align:center"><%=vUnit_L1%></td>
        <td style="text-align:center"><%=vUnit_L1Title%></td>
      </tr>
      <tr>
        <td class="rowshade">&nbsp;<%=Session("Completion_L0Tit")%> ID</td>
        <td class="rowshade"><%=Session("Completion_L0Tit")%> Name</td>
      </tr>
      <tr>
        <td style="text-align:center"><input type="text" name="vUnit_L0" size="8" value="<%=vUnit_L0%>"></td>
        <td style="text-align:center"><input type="text" name="vUnit_L0Title" size="40" value="<%=vUnit_L0Title%>"></td>
        <td style="text-align:center"><input type="submit" value="Add" name="vReg" id="vReg" class="button"></td>
      </tr>
    </table>
    <input type="hidden" name="vUnit_L1" value="<%=vUnit_L1%>">
    <input type="hidden" name="vUnit_L1Title" value="<%=vUnit_L1Title%>">

  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>


