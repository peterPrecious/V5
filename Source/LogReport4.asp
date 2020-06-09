<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->
<!--#include virtual = "V5/Inc/Db_Parm.asp"-->

<html>

<head>
  <title>LogReport4</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  
  <style>
    .table .t1 th:nth-child(1), .table .t1 td:nth-child(1) { width:15%; text-align:left; }
    .table .t1 th:nth-child(2), .table .t1 td:nth-child(2) { width:15%; text-align:left; }
    .table .t1 th:nth-child(3), .table .t1 td:nth-child(3) { width:15%; text-align:left; }
    .table .t1 th:nth-child(4), .table .t1 td:nth-child(4) { width:15%; text-align:left;  }
    .table .t1 th:nth-child(5), .table .t1 td:nth-child(5) { width:10%; text-align:center; }
    .table .t1 th:nth-child(6), .table .t1 td:nth-child(6) { width:10%; text-align:center; }
    .table .t1 th:nth-child(7), .table .t1 td:nth-child(7) { width:20%; text-align:left; }
  </style>
</head>

<body>

  <% 
    Server.Execute vShellHi

    Server.ScriptTimeout = 60 * 10

    Dim vDays, vLearners, vSelect, vDateFormat, vDate, vType, vCriteria, aCrit, bOk, vPrevious, vCnt, vParmNo, vUrl

    vCriteria   = Replace(fDefault(Request("vCriteria"), "0"), " ", "")
    vDays       = fDefault(Request("vDays"), "30")
    vLearners   = fDefault(Trim(Request("vLearners")), "2")
    vDateFormat = fDefault(Request("vDateFormat"), "X")
    vSelect     = Ucase(Trim(Request("vSelect")))
    vPrevious   = Request("vPrevious")
    vCnt        = Clng(fDefault(Request("vCnt"), 0))
    vParmNo     = Request("vParmNo")
  
    '...Excel
    If Request.Form("bExcel").Count = 1 Then
      vUrl = "LogReport4_x.asp"_
           & "?vCriteria="   & Server.UrlEncode(vCriteria) _
           & "&vDays="       & Server.UrlEncode(vDays) _
           & "&vLearners="   & Server.UrlEncode(vLearners) _
           & "&vDateFormat=" & Server.UrlEncode(vDateFormat) _
           & "&vSelect="     & Server.UrlEncode(vSelect) _
           & "&vPrevious="   & Server.UrlEncode(vPrevious) _
           & "&vParmNo="     & Server.UrlEncode(vParmNo) _
           & "&vCriteria="   & Server.UrlEncode(vCriteria)
         
      Response.Redirect vUrl
    End If

    If Request.Form("vHidden").Count = 0 and vCnt = 0 Then

  %>

  <h1>Completion Report</h1>
  <h2>This report displays Active Learners, sorted by Last Name, where the selected Learning Modules have been completed during the selected period.</h2>
  <br />

  <form method="POST" action="LogReport4.asp">
    <input type="Hidden" name="vHidden" value="Hidden">
    <input type="Hidden" name="vParmNo" value="<%=vParmNo%>">
    <table class="table">
      <tr>
        <th align="right" nowrap width="35%" valign="top">Completed within the last :</th>
        <td width="65%"><select size="1" name="vDays">
        <option value="7"   <%=fselect("7",   vdays)%>>7</option>
        <option value="30"  <%=fselect("30",  vdays)%>>30</option>
        <option value="60"  <%=fselect("60",  vdays)%>>60</option>
        <option value="90"  <%=fselect("90",  vdays)%>>90</option>
        <option value="180" <%=fselect("180", vdays)%>>180</option>
        <option value="365" <%=fselect("365", vdays)%>>365</option>
        <option value="0"   <%=fselect("0",   vdays)%>>All</option>
        </select> days</td>
      </tr>
      <tr>
        <th align="right" nowrap width="35%" valign="top">Include :</th>
        <td width="65%"> 
          <input type="checkbox" name="vLearners" value="2" <%=fchecks(vlearners, "2")%>>Learners<br>
          <input type="checkbox" name="vLearners" value="3" <%=fchecks(vlearners, "3")%>>Facilitators<br>
          <input type="checkbox" name="vLearners" value="4" <%=fchecks(vlearners, "4")%>>Managers<br>
          <input type="checkbox" name="vLearners" value="5" <%=fchecks(vlearners, "5")%>>Administrators 
        </td>
      </tr>
      <% 
        i = fCriteriaList (svCustAcctId, "REPT:" & svMembCriteria)
        If vCriteriaListCnt > 1 Then
      %>
      <tr>
        <th align="right" width="35%" valign="top">from Group :</th>
        <td width="65%"><select size="<%=vCriteriaListCnt%>" name="vCriteria" multiple><%=i%></select></td>
      </tr>
      <%  
        Else 
      %>
      <input type="hidden" name="vCriteria" value="<%=svMembCriteria%>">
      <tr>
        <th align="right" width="35%" height="30">from Group :</th>
        <td width="65%" height="30"><%=fCriteria (svMembCriteria)%></td>
      </tr>
      <% 
        End If 
      %>
      <tr>
        <th align="right" nowrap width="35%" valign="top">Select Modules :</th>
        <td width="65%">
          <input type="text" name="vSelect" size="52" value="<%=vSelect%>"><br>Leave empty to include all designated Modules or <br>enter Modules Ids separated by spaces, ie &quot;1234EN 3456EN&quot;.</td>
      </tr>
      <tr>
        <th align="right" width="35%" valign="top">Date Format : </th>
        <td width="65%">
          <input type="radio" name="vDateFormat" value="S" <%=fCheck("S", vDateFormat)%>>MMM DD, YYYY (ie Jan 31, 2010)<br>
          <input type="radio" name="vDateFormat" value="X" <%=fCheck("X", vDateFormat)%>>MM/DD/YYYY (ie 01/31/2010)</td>
      </tr>
      </table>

      <div style="text-align:center; margin:20px; padding:20px;">
        <input type="submit" value="Online" name="bPrint" id="bPrint" class="button070">&ensp; or &ensp; 
        <input type="submit" value="Excel" name="bExcel" class="button070">
      </div>


  </form>
  <%
    Else
  %>

  <h1>Completion Report</h1>
  <h2>This report displays Active Learners, sorted by Last Name, who completed their Learning Modules during the period selected.</h2>
  <br />

  <table class="table">
    <tr class="t1">
      <th class="rowshade">Group</th>
      <th class="rowshade">First Name </th>
      <th class="rowshade">Last Name</th>
      <th class="rowshade"><%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%></th>
      <th class="rowshade">Date Completed</th>
      <th class="rowshade">Module</th>
      <th class="rowshade">Title</th>
    </tr>
    <%
      vSql = "SELECT "_
      	   & "  Memb.Memb_Id, "_
      	   & "  Memb.Memb_LastName, "_
      	   & "  Memb.Memb_FirstName, "_
      	   & "  Memb.Memb_Criteria, "_
      	   & "  Memb.Memb_Level, "_
      	   & "  Logs.Logs_Posted as Posted, "_ 
      	   & "  CASE Logs_Type WHEN 'S' THEN Logs.Logs_Item WHEN 'L' THEN LEFT(Logs.Logs_Item, 6) END AS Id "_ 

      	   & "FROM "_
      	   & "  Memb INNER JOIN "_
      	   & "  Logs ON Memb.Memb_No = Logs.Logs_MembNo "_

      	   & "WHERE "_

      	   & "  (Logs.Logs_AcctId = '" & svCustAcctId & "') "_
      	   & "  AND (Memb.Memb_LastName + Memb.Memb_FirstName + Memb.Memb_Id >= '" & vPrevious & "') "_
      	   & "  AND (Memb.Memb_Active = 1) "_
      	   & "  AND (Memb.Memb_Level IN (" & vLearners & ")) "_
           &    fIf(vDays = 0, "", "AND (Logs.Logs_Posted >='" & fFormatSqlDate(DateAdd("d", Now(), -vDays)) & "') ") _
      	   & "  AND (Logs.Logs_Type = 'S' AND LEN(Logs.Logs_Item) = 6) "_
           &    fParmValue (vParmNo) _
      	   
      	   & "OR "_
      	   & "  (Logs.Logs_AcctId = '" & svCustAcctId & "') "_
      	   & "  AND (Memb.Memb_LastName + Memb.Memb_FirstName + Memb.Memb_Id >= '" & vPrevious & "') "_
      	   & "  AND (Memb.Memb_Active = 1) "_
      	   & "  AND (Memb.Memb_Level IN (" & vLearners & ")) "_
           &    fIf(vDays = 0, "", "AND (Logs.Logs_Posted >='" & fFormatSqlDate(DateAdd("d", Now(), -vDays)) & "') ") _
      	   & "  AND (Logs.Logs_Type = 'L' AND Logs.Logs_Item LIKE '%_completed') "_
           &    fParmValue (vParmNo) _

      	   & "ORDER BY "_
      	   & "  Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_Id "


'    sDebug
      sOpenDb
      Set oRs = oDb.Execute(vSql)
      vCnt = 0

      Do While Not oRs.Eof
        bOk = False
        If vSelect = "" Or Instr(vSelect, oRs("Id")) > 0 Then 
          bOk = True
        End If

        '...criteria can be 129,330 or just 129
        If bOk Then
          If Len(vCriteria) > 3 Then          
            bOk = False
            aCrit = Split(vCriteria, ",")
            For i = 0 To Ubound(aCrit)
              If Instr(oRs("Memb_Criteria"), aCrit(i)) > 0 Then
                bOk = True
                Exit For
              End If
            Next    
          End If   
        End If

        If bOk Then
          vCnt = vCnt + 1
          If Clng(vCnt) Mod 100 = 0 Then 
            vPrevious = oRs("Memb_LastName") & oRs("Memb_FirstName") & oRs("Memb_Id") 
          	Exit Do
          End If
    %>
    <tr class="t1">
      <td><%=fIf(Len(Trim(oRs("Memb_Criteria"))) < 3 Or Trim(oRs("Memb_Criteria")) = "0" , "", Replace(fCriteria(oRs("Memb_Criteria")), " + ", "<br>"))%></td>
      <td><%=oRs("Memb_FirstName")%></td>
      <td><%=oRs("Memb_LastName")%></td>
      <td><%=fIf(oRs("Memb_Level")=5, "********", oRs("Memb_Id"))%> </td>
      <td>
      <% 
         vDate = oRs("Posted")
      	 If vDateFormat = "S" Then 
           vDate = fFormatDate(vDate)
         Else
           vDate = Right("00" & Month(vDate), 2) & "/" & Right("00" & Day(vDate), 2) & "/" & Year(vDate)
         End If 
      %>
      <%=vDate%>
      </td>
      <td><%=oRs("Id")%> </td>
      <td><%=fLeft(fModsTitle(oRs("Id")),64)%> </td>
    </tr>
    <%
        End If
        oRs.MoveNext	        
      Loop
      sCloseDB
    %>
  </table>


  <div style="text-align:center; padding:20px; margin:20px;">
    <input onclick="location.href = 'LogReport4.asp?vCnt=0&vPrevious=<%=vPrevious%>&vDays=<%=vDays%>&vCriteria=<%=vCriteria%>&vLearners=<%=vLearners%>&vDateFormat=<%=vDateFormat%>&vSelect=<%=Replace(vSelect, " ", "+")%>&vParmNo=<%=vParmNo%>'" type="button" value="Return" name="bNext" class="button085">
    <% If Clng(vCnt) > 0 And Clng(vCnt) Mod 100 = 0 Then '...If next group, get next starting value %>
    <input onclick="location.href = 'LogReport4.asp?vCnt=<%=vCnt%>&vPrevious=<%=vPrevious%>&vDays=<%=vDays%>&vCriteria=<%=vCriteria%>&vLearners=<%=vLearners%>&vDateFormat=<%=vDateFormat%>&vSelect=<%=Replace(vSelect, " ", "+")%>&vParmNo=<%=vParmNo%>'" type="button" value="Next" name="bNext" class="button085"></p>
    <% End If %>
  </div>

  <%
	  End If
  %>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>