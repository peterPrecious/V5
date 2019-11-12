<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Jobs.asp"-->
<!--#include file = "Completion_Routines.asp"-->

<%
  Dim vUrl, vActive, vNext, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindCriteria, vFormat, vLearners, vChoice
  Dim vMyGr, vMyL1, vMyL0, vMyRl, vMyCh, vWhere, vSelect, vCnt

  '...First time in?  
  If Session("Completion_InitParms") = "" Then 
    Session("Completion_InitParms") = "Y"
    Response.Redirect "Completion.asp?vNext=Completion_Learners.asp" 
  End If

  '...determine rights of user (RRRR SSSS R)
  vMyGr = fCriteria (svMembCriteria)
  vMyL1 = Left(vMyGr, Session("Completion_L1len"))
  vMyL0 = Mid(vMyGr, Session("Completion_L0str"), Session("Completion_L0len"))
  vMyRl = Right(vMyGr, Session("Completion_RLlen"))
	vMyCh = fRole_Children(vMyRL)

  vNext            = Request("vNext")
  vActive          = fDefault(Request("vActive"), "1")
  vFind            = fDefault(Request("vFind"), "S")
  vFindId          = fUnQuote(Request("vFindId"))
  vFindFirstName   = fUnQuote(Request("vFindFirstName"))
  vFindLastName    = fUnQuote(Request("vFindLastName"))
  vFindEmail       = fNoQuote(Request("vFindEmail"))
  vFindCriteria    = fDefault(Replace(Request("vFindCriteria"), " ", ""), "All")
  vFormat          = fDefault(Request("vFormat"), "o")
  vLearners        = fDefault(Request("vLearners"), "n")

  If Request.Form.Count > 0 Then
    vUrl = ""
    vUrl = vUrl & "?vNext="          & vNext
    vUrl = vUrl & "&vActive="        & vActive
    vUrl = vUrl & "&vFind="          & vFind
    vUrl = vUrl & "&vFindId="        & vFindId
    vUrl = vUrl & "&vFindFirstName=" & vFindFirstName
    vUrl = vUrl & "&vFindLastName="  & vFindLastName
    vUrl = vUrl & "&vFindEmail="     & vFindEmail
    vUrl = vUrl & "&vFindCriteria="  & vFindCriteria
    vUrl = vUrl & "&vFormat="        & vFormat
    vUrl = vUrl & "&vLearners="      & vLearners
    Response.Redirect "Completion_Learners_" & vFormat & ".asp" & vUrl
  End If 

%>
<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <script>
  function validate(theForm) {
    <% If Session("Completion_Level") < 4 Then %>

    if (theForm.vFindCriteria.selectedIndex < 0)
    {
      alert("Please select a Location.");
      theForm.vFindCriteria.focus();
      return (false);
    }

    <% End If %>

    return (true);
  }
  </script>



</head>

<body>

  <% Server.Execute vShellHi %>

  <table class="table">

    <form method="POST" action="Completion_Learners.asp" id="fForm" onsubmit="return validate(this)">
      <tr>
        <td colspan="2" bordercolor="#FFFFFF">
        <h1 align="center">Learner Report</h1>
        </td>
      </tr>
      <tr>
        <th>
        Select Learners whose&nbsp;Id's&nbsp;&nbsp; <br />start with :</th>
        <td valign="top" bordercolor="#FFFFFF" width="65%"><input type="text" name="vFindId" size="39" value="<%=vFindId%>" maxlength="6"><br>Leave empty to start at the beginning.</td>
      </tr>
      <tr>
        <th>
        or whose Last Name&nbsp;&nbsp; <br />starts with :</th>
        <td valign="top" bordercolor="#FFFFFF" width="65%"><input type="text" name="vFindLastName" size="39" value="<%=vFindLastName%>"><br>
        Ie Smith (leave empty to start at beginning)</td>
      </tr>
      <tr>
        <th align="right" width="35%" valign="top">
        including :</th>
        <td>
          <input type="radio" value="1" <%=fcheck("1", vactive)%> name="vActive">Active Learners only<br>
          <input type="radio" value="*" <%=fcheck("*", vactive)%> name="vActive">Both Active AND Inactive Learners<br>
          <input type="radio" value="0" <%=fcheck("0", vactive)%> name="vActive">Inactive Learners only</td>
      </tr>

			<%
          i = fLocation(vMemb_Criteria)
      %>
      <tr>
        <th>from Location :</th>
        <td>  
          <select name="vFindCriteria"  style="width: 400px" size="<%=fMin(vCnt, 30)%>" multiple><%=i%></select>
          <br>Select multiple Locations by holding down the CTRL key and clicking on your selections.

          <% If Session("Completion_Level") > 3 Then %>
          <br />
          <span style="margin-top:20px; background-color:yellow">If you do NOT select any locations then All will be included.</span>

          <% End If %>



        </td>
      </tr>

      <tr>
        <th>Format :</th>
        <td>
          <input type="radio" name="vFormat" value="o" <%=fcheck("o", vformat)%>>Online<br>
          <input type="radio" name="vFormat" value="x" <%=fcheck("x", vformat)%>>Excel
        </td>
      </tr>

      <tr>
        <td colspan="2" style="text-align:center; padding:40px;">
          <input onclick="<%=fIf(fNoValue(vNext), "javascript:history.back(1)", "history.back(1)")%>" type="button" value="<%="Return"%>" name="bReturn" id="bReturn" class="Button"><%=f10%><input type="submit" value="<%="Continue"%>" name="bContinue" class="button"> 
        </td>
      </tr>
      <input type="hidden" name="vNext" value="<%=vNext%>">
    </form>

  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
  <!--#include file = "Completion_Footer.asp"-->

</body>

</html>

