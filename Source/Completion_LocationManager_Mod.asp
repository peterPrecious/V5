<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include file = "Completion_Routines.asp"-->
<!--#include file = "Completion_LocationManager_Routines.asp"-->

<%
  vUnit_No = Request("vUnit_No")
  sGetUnit vUnit_No

  If Request.Form("bMod").Count = 1 Then 
    vSql = "UPDATE "_
         & "  V5_Comp.dbo.Unit "_
         & "SET " _
         & "  Unit_L0Title   = '" & fUnQuote(Request("vUnit_L0Title")) & "', " _
         & "  Unit_HO        =  " & Request("vUnit_HO")       & " , " _
         & "  Unit_Active    =  " & Request("vUnit_Active")   & "   " _
         & "WHERE "_
         & "  Unit_No        =  " & Request("vUnit_No")
    sCompletion_Debug
    sOpenDb2
    oDb2.Execute(vSql)

    For Each vFld In Request.Form
      If Left(vFld, 5) = "Role_" Then
        vRole = Mid(vFld, 6)
        vJobs = Replace(Request(vFld).Item, ",", "")
        vSql = " UPDATE Crit SET Crit_JobsId = '" & vJobs & "'" _ 
             & " WHERE Crit_AcctId = '" & svCustAcctId & "' AND Crit_Id = '" & vUnit_L1 & " " & vUnit_L0 & " " & vRole & "' AND ISNULL(Crit_JobsId,'') <> '" & vJobs & "'"
        sCompletion_Debug
        oDb2.Execute(vSql)
      End iF
    Next  
    sCloseDb2    

    Response.Redirect "Completion_LocationManager_Rev.asp?vUnit_No=" & vUnit_No

  End If  
  
%>

<html>

<head>
  <title>Completion_LocationManager_Mod</title>
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

    function validateLoc (theForm) {
      if (theForm.vUnit_L0.value.length != L0len) {
        alert("Please enter a " + vL0len + " character " + vL0tit + " Id.");
        theForm.vUnit_L0.focus();
        return (false);
      }
      if (theForm.vUnit_L0Title.value == "") {
        alert("Please enter the " + vL0tit + " Name.");
        theForm.vUnit_L0Title.focus();
        return (false);
      }    
      if (theForm.vUnit_L0Title.value.length < 4 || theForm.vUnit_L0Title.value.length > 128) {
        alert("The " + vL0tit + " Name must be between 4 and 128 characters.");
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
    <h1><%=Session("Completion_L0Tit")%> List (Edit)</h1>
    <h2 style="text-align: left;">Edit the fields below and Select the appropriate Training Plans.&nbsp; Note: ensure you hold down the CTRL button as you click if you wish to select multiple training plans.&nbsp; Click <b>Update</b> when finished.</h2>
  </div>

  <form method="POST" id="fLoc" name="fLoc" action="Completion_LocationManager_Mod.asp" onsubmit="return validateLoc(this)">

    <input type="hidden" name="vUnit_No" value="<%=vUnit_No%>">
    <input type="hidden" name="vUnit_L1" value="<%=vUnit_L1%>">
    <input type="hidden" name="vUnit_L1Title" value="<%=vUnit_L1Title%>">

    <table style="width: 650px; margin: auto;">

      <tr>
        <td class="rowshade" style="width:150px"><%=Session("Completion_L1Tit")%></td>
        <td class="rowshade" style="width:1500px"><%=Session("Completion_L0Tit")%> ID</td>
        <td class="rowshade" style="width:200px"><%=Session("Completion_L0Tit")%> Name</td>
        <td class="rowshade" style="width:150px">Head Office?</td>
        <td class="rowshade" style="width:150px">Active?</td>
      </tr>

      <tr>
        <td style="text-align: center"><b><%=vUnit_L1 & " (" & vUnit_L1Title & ")"%></b> </td>
        <td style="text-align: center">
          <% If Len(vUnit_L1) > 0 And Len(vUnit_L0) > 0 Then %> <b><%=vUnit_L0%></b>
          <input type="hidden" name="vUnit_L0" value="<%=vUnit_L0%>"><% Else %>
          <input type="text" size="6" name="vUnit_L0" maxlength="4">
          <% End If %> 
        </td>
        <td style="text-align: center">
          <input type="text" name="vUnit_L0Title" maxlength="128" value="<%=vUnit_L0Title%>">
        </td>
        <td style="text-align: center">
          <input type="radio" value="1" name="vUnit_HO" <%=fcheck("1", vunit_ho)%>>Yes
          <input type="radio" value="0" name="vUnit_HO" <%=fcheck("0", vunit_ho)%>>No 
        </td>
        <td style="text-align: center">
          <input type="radio" value="1" name="vUnit_Active" <%=fcheck("1", vunit_active)%>>Yes
          <input type="radio" value="0" name="vUnit_Active" <%=fcheck("0", vunit_active)%>>No 
        </td>
      </tr>
      <tr>
        <th style="text-align: center">&nbsp;</th>
        <th colspan="4" style="text-align: center">&nbsp;</th>
      </tr>
      <tr>
        <td class="rowshade">Role</td>
        <th colspan="4" class="rowshade">&nbsp;Training Plan</th>
      </tr>
      <% 
        Dim vRoleLen, vLocnLen
        i = ""
        vRoleLen = Session("Completion_RLlen")
        vLocnLen = Session("Completion_L1len") + Session("Completion_L0len") + 1
        vSql = ""_
             & " SELECT"_     
             & "   RIGHT(Crit.Crit_Id, " & Session("Completion_RLlen") & ") AS Role, Crit.Crit_JobsId AS Jobs"_
             & "  FROM"_         
             & "    V5_Comp.dbo.Unit Unit WITH (NOLOCK) INNER JOIN"_
             & "    V5_Vubz.dbo.Crit Crit WITH (NOLOCK) ON Unit.Unit_L1 + ' ' + Unit.Unit_L0 = LEFT(Crit.Crit_Id, " & vLocnLen & ") AND Unit.Unit_AcctId = Crit.Crit_AcctId"_
             & "  WHERE"_     
             & "    (Unit.Unit_No = " & Request("vUnit_No") & ")"
        sCompletion_Debug
        sOpenDb
        Set oRs = oDb.Execute(vSql)
        Do While Not oRs.Eof
      %>
      <tr>
        <th style="text-align: center"><%=oRs("Role")%></th>
        <td colspan="4" style="text-align: center">
          <select size="12" name="Role_<%=oRs("Role")%>" style="width: 350px" multiple><%= fJobs (oRs("Jobs"))%></select></td>
      </tr>
      <%    
        oRs.MoveNext
        Loop
        Set oRs = Nothing
        sCloseDb
      %>
      <tr>
        <td style="text-align: center" colspan="5" height="50">
          <input onclick="location.href = 'Completion_LocationManager.asp'" type="button" value="Cancel" name="bCancel" class="button070">
          <%=f10%>
          <input type="submit" value="Update" name="bMod" class="button070">
        </td>
      </tr>
    </table>

  </form>


  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>
</html>
