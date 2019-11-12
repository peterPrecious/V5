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

  <h1><!--webbot bot='PurpleText' PREVIEW='Activity Report'--><%=fPhra(000487)%></h1>
  <h2><!--webbot bot='PurpleText' PREVIEW='This report, sorted by Last Name, shows the Time Spent in minutes reviewing Modules and any Scores achieved in Assessments.'--><%=fPhra(000550)%></h2>
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
              Case 1 : j =  1   : vDesc = fPhraH(000274)
              Case 2 : j =  7   : vDesc = "7 " & fPhraH(000115)
              Case 3 : j = 14   : vDesc = j & " " & fPhraH(000115)
              Case 4 : j = 30   : vDesc = j & " " & fPhraH(000115)
              Case 5 : j = 60   : vDesc = j & " " & fPhraH(000115)
              Case 6 : j = 90   : vDesc = j & " " & fPhraH(000115)
              Case 7 : j = 180  : vDesc = j & " " & fPhraH(000115)
              Case 8 : j = 365  : vDesc = j & " " & fPhraH(000115)
              Case 9 : j = 9999 : vDesc = fPhraH(000340)
            End Select
            k = fFormatSqlDate(DateAdd ("d", -j, Now))
            If j = 9999 Then k = "Jan 1, 2000"
            vSelected = fIf(vStrDate = k, " selected", "")
            vOption = vOption & "<option value='" & k & "'" & vSelected & ">" & vDesc & "</option>" & vbCrLf 
          Next
        %> 
        <!--webbot bot='PurpleText' PREVIEW='For the last'--><%=fPhra(001671)%> :</th>
        <td><select size="1" name="vStrDate"><%=vOption%></select></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Include'--><%=fPhra(000155)%> :</th>
        <td>
          <input type="radio" name="vActive" value="y" <%=fcheck("y", vActive)%>> <!--webbot bot='PurpleText' PREVIEW='Active Learners'--><%=fPhra(001672)%><br />
          <input type="radio" name="vActive" value="n" <%=fcheck("n", vActive)%>><!--webbot bot='PurpleText' PREVIEW='All Learners'--><%=fPhra(001673)%>
        </td>
      </tr>



<% If Len(vModsOnly) = 0 Then %>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Only report on modules'--><%=fPhra(000557)%> :</th>
        <td><input type="text" name="vMods" size="29" value="<%=vMods%>"> <!--webbot bot='PurpleText' PREVIEW='ie 1234EN 1234FR.'--><%=fPhra(000561)%><br><!--webbot bot='PurpleText' PREVIEW='Separate Module IDs with a space.&nbsp; Leave empty to select all modules.'--><%=fPhra(000558)%></td>
      </tr>
<% Else %>
      <input type="hidden" value="<%=vMods%>" name="vMods">
      <tr>
        <th height="30"><!--webbot bot='PurpleText' PREVIEW='for modules'--><%=fPhra(000566)%> :</th>
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
        <th><!--webbot bot='PurpleText' PREVIEW='Find learners that'--><%=fPhra(000541)%> :</th>
        <td><input type="radio" name="vFind" value="S" <%=fcheck("s", vfind)%>><!--webbot bot='PurpleText' PREVIEW='start with'--><%=fPhra(000463)%> or <input type="radio" name="vFind" value="C" <%=fcheck("c", vfind)%>><!--webbot bot='PurpleText' PREVIEW='contain'--><%=fPhra(000464)%></td>
      </tr>
      <tr>
        <th>&nbsp;<%=fIf(svCustPwd, fPhraH(000411), fPhraH(000211))%> : </th>
        <td>&nbsp;&nbsp;&nbsp; <input type="text" name="vFindId" size="29" value="<%=vFindId%>"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%> : </th>
        <td>&nbsp;&nbsp;&nbsp; <input type="text" name="vFindFirstName" size="29" value="<%=vFindFirstName%>">&nbsp; </td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%> :</th>
        <td>&nbsp;&nbsp;&nbsp; <input type="text" name="vFindLastName" size="29" value="<%=vFindLastName%>"></td>
      </tr>
      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Email Address'--><%=fPhra(000126)%> :</th>
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
        <th><!--webbot bot='PurpleText' PREVIEW='from Group'--><%=fPhra(000565)%> :</th>
        <td>&nbsp;&nbsp;&nbsp; <select size="<%=vCriteriaListCnt%>" name="vFindCriteria" multiple><%=i%></select></td>
      </tr>
      <%  
          Else 
      %>
      <input type="hidden" name="vFindCriteria" value="<%=svMembCriteria%>">
      <tr>
        <th>
        <!--webbot bot='PurpleText' PREVIEW='from Group'--><%=fPhra(000565)%> :</th>
        <td><%=fCriteria (svMembCriteria)%>&nbsp;&nbsp;&nbsp; </td>
      </tr>
      <% 
        End If 
      %>

      <tr>
        <th><!--webbot bot='PurpleText' PREVIEW='Output'--><%=fPhra(001674)%> : </th>
        <td>
          <input type="radio" name="vFormat" value="o" <%=fcheck("o", vformat)%>><!--webbot bot='PurpleText' PREVIEW='Online'--><%=fPhra(000488)%> 
          <input type="radio" name="vFormat" value="x" <%=fcheck("x", vformat)%>><!--webbot bot='PurpleText' PREVIEW='Excel'--><%=fPhra(000560)%>
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

