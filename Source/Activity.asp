<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<% 
  Dim vNext, vCurList, vStrDate, vActive, vMods, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vFormat
  Dim vModsOnly '...this is when we receive a requrest from outside the app to select only certain mods

  vCurList       = fDefault(Request("vCurList"), 0)
  vStrDate       = fDefault(Request("vStrDate"), fFormatSqlDate(DateAdd ("d", -90, Now)))
  vActive        = fDefault(Request("vActive"), "y")
  vModsOnly      = Request("vModsOnly")
  If Len(vModsOnly) > 0 Then
    vMods        = vModsOnly '...format: Activity.asp?vModsOnly=1234EN,1235EN,1440FR
  Else
    vMods        = Request("vMods")
  End If
  vFind          = fDefault(Request("vFind"), "S")
  vFindId        = fUnQuote(Request("vFindId"))
  vFindFirstName = fUnQuote(Request("vFindFirstName"))
  vFindLastName  = fUnQuote(Request("vFindLastName"))
  vFindEmail     = fNoQuote(Request("vFindEmail"))
  vFindMemo      = fUnQuote(Request("vFindMemo"))
  vFindCriteria  = fDefault(Request("vFindCriteria"), "0")
  vFormat        = fDefault(Request("vFormat"), "o")


  '...processing the form?
  If Request("vForm").Count > 0 Then
    Session("soRs") = "" 
  
    '...goto online or excel reports
    vNext = "Activity_" & vFormat  & ".asp"     _
          & "?vStrDate="       & vStrDate       _
          & "&vCurList="       & vCurList       _
          & "&vActive="        & vActive        _
          & "&vMods="          & vMods          _
          & "&vModsOnly="      & vModsOnly      _
          & "&vFind="          & vFind          _
          & "&vFindId="        & vFindId        _
          & "&vFindFirstName=" & vFindFirstName _
          & "&vFindLastName="  & vFindLastName  _
          & "&vFindEmail="     & vFindEmail     _
          & "&vFindMemo="      & vFindMemo      _
          & "&vFindCriteria="  & vFindCriteria  
    Response.Redirect vNext
'   Response.Write vParm
  End If
  
%>

<html>

<head>
  <title>Activity</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

  <% Server.Execute vShellHi %>

  <h1><!--[[-->Activity Report<!--]]--></h1>
  <h2><!--[[-->This report, sorted by Last Name, shows the Time Spent in minutes reviewing Modules and any Scores achieved in Assessments.<!--]]--></h2>
  <br /><br />
  <form method="POST" action="Activity.asp">
    <table class="table">
      <tr>
        <th>
        <%    
          Dim vOption, vDesc, vSelected
          vOption = ""
          vSelected = ""
          For i = 1 To 9
            Select Case i
              Case 1 : j =  1   : vDesc = "<!--{{-->1 day<!--}}-->"
              Case 2 : j =  7   : vDesc = "7 " & "<!--{{-->days<!--}}-->"
              Case 3 : j = 14   : vDesc = j & " " & "<!--{{-->days<!--}}-->"
              Case 4 : j = 30   : vDesc = j & " " & "<!--{{-->days<!--}}-->"
              Case 5 : j = 60   : vDesc = j & " " & "<!--{{-->days<!--}}-->"
              Case 6 : j = 90   : vDesc = j & " " & "<!--{{-->days<!--}}-->"
              Case 7 : j = 180  : vDesc = j & " " & "<!--{{-->days<!--}}-->"
              Case 8 : j = 365  : vDesc = j & " " & "<!--{{-->days<!--}}-->"
              Case 9 : j = 9999 : vDesc = "<!--{{-->all available days<!--}}-->"
            End Select
            k = fFormatSqlDate(DateAdd ("d", -j, Now))
            If j = 9999 Then k = "Jan 1, 2000"
            vSelected = fIf(vStrDate = k, " selected", "")
            vOption = vOption & "<option value='" & k & "'" & vSelected & ">" & vDesc & "</option>" & vbCrLf 
          Next
        %> 
        <!--[[-->For the last<!--]]--> :</th>
        <td><select size="1" name="vStrDate"><%=vOption%></select></td>
      </tr>
      <tr>
        <th><!--[[-->Include<!--]]--> :</th>
        <td>
          <input type="radio" name="vActive" value="y" <%=fcheck("y", vActive)%>> <!--[[-->Active Learners<!--]]--><br />
          <input type="radio" name="vActive" value="n" <%=fcheck("n", vActive)%>><!--[[-->All Learners<!--]]-->
        </td>
      </tr>



<% If Len(vModsOnly) = 0 Then %>
      <tr>
        <th><!--[[-->Only report on modules<!--]]--> :</th>
        <td><input type="text" name="vMods" size="29" value="<%=vMods%>"> <!--[[-->ie 1234EN 1234FR.<!--]]--><br><!--[[-->Separate Module IDs with a space.&nbsp; Leave empty to select all modules.<!--]]--></td>
      </tr>
<% Else %>
      <input type="hidden" value="<%=vMods%>" name="vMods">
      <tr>
        <th height="30"><!--[[-->for modules<!--]]--> :</th>
        <td height="30">
        <%
            Dim aMods
            aMods = Split(vMods, ",")
            For i = 0 To Ubound(aMods)                 
              Response.Write fIf(i = 0,"","<br>") & aMods(i) & " - " & fModsTitle(aMods(i))
            Next
        %>
        </td>
      </tr>
<% End If %>
      <tr>
        <th><!--[[-->Find learners that<!--]]--> :</th>
        <td><input type="radio" name="vFind" value="S" <%=fcheck("s", vfind)%>><!--[[-->start with<!--]]--> or <input type="radio" name="vFind" value="C" <%=fcheck("c", vfind)%>><!--[[-->contain<!--]]--></td>
      </tr>
      <tr>
        <th>&nbsp;<%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%> : </th>
        <td>&nbsp;&nbsp;&nbsp; <input type="text" name="vFindId" size="29" value="<%=vFindId%>"></td>
      </tr>
      <tr>
        <th><!--[[-->First Name<!--]]--> : </th>
        <td>&nbsp;&nbsp;&nbsp; <input type="text" name="vFindFirstName" size="29" value="<%=vFindFirstName%>">&nbsp; </td>
      </tr>
      <tr>
        <th><!--[[-->Last Name<!--]]--> :</th>
        <td>&nbsp;&nbsp;&nbsp; <input type="text" name="vFindLastName" size="29" value="<%=vFindLastName%>"></td>
      </tr>
      <tr>
        <th><!--[[-->Email Address<!--]]--> :</th>
        <td>&nbsp;&nbsp;&nbsp; <input type="text" name="vFindEmail" size="29" value="<%=vFindEmail%>"></td>
      </tr>
      <tr>
        <th>Memo :</th>
        <td>&nbsp;&nbsp;&nbsp; <input type="text" name="vFindMemo" size="29" value="<%=vFindMemo%>"></td>
      </tr>

      <% 
        i = fCriteriaList (svCustAcctId, "REPT:" & svMembCriteria)
        If vCriteriaListCnt > 1 Then
      %>
      <tr>
        <th><!--[[-->from Group<!--]]--> :</th>
        <td>&nbsp;&nbsp;&nbsp; <select size="<%=vCriteriaListCnt%>" name="vFindCriteria" multiple><%=i%></select></td>
      </tr>
      <%  
          Else 
      %>
      <input type="hidden" name="vFindCriteria" value="<%=svMembCriteria%>">
      <tr>
        <th>
        <!--[[-->from Group<!--]]--> :</th>
        <td><%=fCriteria (svMembCriteria)%>&nbsp;&nbsp;&nbsp; </td>
      </tr>
      <% 
        End If 
      %>

      <tr>
        <th><!--[[-->Output<!--]]--> : </th>
        <td>
          <input type="radio" name="vFormat" value="o" <%=fcheck("o", vformat)%>><!--[[-->Online<!--]]--> 
          <input type="radio" name="vFormat" value="x" <%=fcheck("x", vformat)%>><!--[[-->Excel<!--]]-->
        </td>
      </tr>

      <tr>
        <td colspan="2" style="text-align:center; padding:20px;">
          <input type="submit" value="<%=bContinue%>" name="bContinue" class="button"> 
          <input type="hidden" value="<%=Request("vParmNo")%>" name="vParmNo">
        </td>
      </tr>

    </table>

    <input type="hidden" name="vForm" value="y">
    <input type="hidden" name="vModsOnly" value="<%=vModsOnly%>">

  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>