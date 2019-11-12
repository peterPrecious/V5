<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_ExamReport.asp"-->

<html>

<head>
  <title>Activity_O</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <style>

    .ellipsis { width:200px; overflow: hidden; white-space: nowrap; text-overflow: ellipsis; display: block; -o-text-overflow: ellipsis; }

    .table tr th:nth-child(01) { width: 10%; text-align: left; }
    .table tr th:nth-child(02) { width: 15%; text-align: left; }
    .table tr th:nth-child(03) { width: 10%; text-align: left; }
    .table tr th:nth-child(04) { width: 10%; text-align: center; }
    .table tr th:nth-child(05) { width: 10%; text-align: center; }
    .table tr th:nth-child(06) { width: 15%; text-align: left; }
    .table tr th:nth-child(07) { width: 10%; text-align: center; white-space: nowrap; }
    .table tr th:nth-child(08) { width: 10%; text-align: center; }
    .table tr th:nth-child(09) { width: 10%; text-align: center; }

    .table tr td:nth-child(01) { width: 10%; text-align: left; }
    .table tr td:nth-child(02) { width: 15%; text-align: left; }
    .table tr td:nth-child(03) { width: 10%; text-align: left; }
    .table tr td:nth-child(04) { width: 10%; text-align: center; white-space: nowrap; }
    .table tr td:nth-child(05) { width: 10%; text-align: center; white-space: nowrap; }
    .table tr td:nth-child(06) { width: 15%; overflow: hidden; }
    .table tr td:nth-child(07) { width: 10%; text-align: center; white-space: nowrap; }
    .table tr td:nth-child(08) { width: 10%; text-align: center; }
    .table tr td:nth-child(09) { width: 10%; text-align: center; white-space: nowrap; }

    .table tr:nth-child(odd) { background-color: #F2F9FD; }
    .table tr:hover { background-color: yellow; }

  </style>

</head>

<body>

  <% Server.Execute vShellHi %>

  <h1><!--webbot bot='PurpleText' PREVIEW='Activity Report'--><%=fPhra(000487)%></h1>
  <h2><!--webbot bot='PurpleText' PREVIEW='This report, sorted by Last Name, shows the Time Spent in minutes reviewing Modules and any Scores achieved in Assessments.'--><%=fPhra(000550)%></h2>
  <br /><br />

  <table class="table">
    <tr>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='Group'--><%=fPhra(000369)%></th>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='Name'--><%=fPhra(000187)%></th>
      <th class="rowshade"><%=fIf(svCustPwd, fPhraH(000411), fPhraH(000211))%></th>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='Active?'--><%=fPhra(000551)%></th>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='Module'--><%=fPhra(000272)%></th>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='Title'--><%=fPhra(000019)%></th>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='Time Spent'--><%=fPhra(000552)%></th>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='Score'--><%=fPhra(000232)%></th>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='Date'--><%=fPhra(000112)%></th>
    </tr>
    <% 

      Dim vNext, vActive, vMods, vModsOnly, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vFormat, vStrDate
      Dim vMembIdLast, vBg, vTitle, vScore, vTimeSpent, vPassword
      Dim vLogsType, vLogsModule, vLogsValue, vLogsTimespent, vLogsPosted          

      vCurList       = Request("vCurList") 
      vMaxList       = fDefault(Request("vMaxList"), 50)
      vStrDate       = Request("vStrDate") 
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
      vFindCriteria  = Request("vFindCriteria")

      '...Get initial recordset on first pass and store in session variable
      If vCurList = 0 Then 

        vSql = "SELECT Memb_Criteria, Memb.Memb_No, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Level, Memb_Active,  " _
             & "CASE Len(Logs.Logs_Item) WHEN 21 THEN SUBSTRING (Logs.Logs_Item,9, 6) ELSE LEFT(Logs.Logs_Item, 6) END AS [Logs_Module], " _
             & "CAST(RIGHT(Logs.Logs_Item, 3) AS FLOAT) AS Logs_Value, " _
             & "SUBSTRING(Logs.Logs_Item, 8, 1) AS Logs_Attempt, " _
             & "Logs.Logs_Posted AS Logs_Posted, " _
             & "CASE LEN(Logs_Item) WHEN 21 THEN 'P' WHEN 10 THEN 'M' WHEN 12 THEN 'E' END AS [Logs_Type], " _
             & "RIGHT(Logs.Logs_Item, 6) AS [Logs_Timespent] " _
             
             & "FROM Memb WITH (nolock) " & fIf(vActive="y", "INNER", "LEFT OUTER") & " JOIN Logs WITH (nolock) " _ 

             & "ON Logs_MembNo = Memb_No " _    
             & "AND (CHARINDEX(Logs.Logs_Type, 'PT') > 0) " _

             & fIf(Len(Trim(vStrDate)) > 0, "AND (Logs.Logs_Posted >= '" & vStrDate & "') ", "" ) _  

             & fIf(Len(Trim(vMods)) > 0, "AND ((LEN(Logs.Logs_Item) = 21) AND (CHARINDEX(SUBSTRING(Logs.Logs_Item, 9, 6), '" & vMods & "') > 0) OR (LEN(Logs.Logs_Item) = 10) AND (CHARINDEX(LEFT(Logs.Logs_Item, 6), '" & vMods & "') > 0)) ", "" )_  

             & "WHERE (Memb_AcctId = '" & svCustAcctId & "') " _
             & "AND (Logs.Logs_Item NOT LIKE '%undefined%') " _             
             & "AND (Logs.Logs_AcctId = '" & svCustAcctId & "') " _             
             & "AND (Memb_Level <= " & svMembLevel & ") " _

             & fIf(vFind = "S" And Len(vFindId) > 0,        "AND (Memb_Id        LIKE '" & vFindId         & "%') ", "" ) _
             & fIf(vFind = "S" And Len(vFindFirstName) > 0, "AND (Memb_FirstName LIKE '" & vFindFirstName  & "%') ", "" ) _
             & fIf(vFind = "S" And Len(vFindLastName) > 0,  "AND (Memb_LastName  LIKE '" & vFindLastName   & "%') ", "" ) _
             & fIf(vFind = "S" And Len(vFindEmail) > 0,     "AND (Memb_Email     LIKE '" & vFindEmail      & "%') ", "" ) _
             & fIf(vFind = "S" And Len(vFindMemo) > 0,      "AND (Memb_Memo      LIKE '" & vFindMemo       & "%') ", "" ) _

             & fIf(vFind = "C" And Len(vFindId) > 0,        "AND (Memb_Id        LIKE '%" & vFindId        & "%') ", "" ) _
             & fIf(vFind = "C" And Len(vFindFirstName) > 0, "AND (Memb_FirstName LIKE '%" & vFindFirstName & "%') ", "" ) _
             & fIf(vFind = "C" And Len(vFindLastName) > 0,  "AND (Memb_LastName  LIKE '%" & vFindLastName  & "%') ", "" ) _
             & fIf(vFind = "C" And Len(vFindEmail) > 0,     "AND (Memb_Email     LIKE '%" & vFindEmail     & "%') ", "" ) _
             & fIf(vFind = "C" And Len(vFindMemo) > 0,      "AND (Memb_Memo      LIKE '%" & vFindMemo      & "%') ", "" ) _

             & fIf(Len(vFindCriteria)> 2,                   "AND (Memb_Criteria = '"      & vFindCriteria  & "')  ", "" ) _

             & "ORDER BY Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName, Memb.Memb_No, " _
             & "CASE Len(Logs.Logs_Item) WHEN 21 THEN Substring(Logs_Item, 9, 6) ELSE LEFT(Logs_Item, 6) END, " _ 
             & "CASE LEN(Logs_Item) WHEN 21 THEN 'P' WHEN 10 THEN 'M' ELSE 'E' END DESC " _ 

						 & ", Logs_Posted DESC"	
  
'       sDebug

        sOpenDB
        Set oRs = oDB.Execute(vSql)
        Set Session("soRs") = oRs
        vCurList = 1
      Else  
        Set oRs = Session("soRs")
      End If  

      vMembIdLast = ""

      '...read until either eof or end of group
      Do While Not oRs.Eof


        vMemb_No        = oRs("Memb_No")
        vMemb_Id        = oRs("Memb_Id")
        vMemb_FirstName = oRs("Memb_FirstName")
        vMemb_LastName  = oRs("Memb_LastName")
        vMemb_Active    = oRs("Memb_Active")
        vMemb_Level     = oRs("Memb_Level")
        vMemb_Criteria  = oRs("Memb_Criteria")

        vLogsType       = fOkValue(oRs("Logs_Type"))
        vLogsModule     = fOkValue(oRs("Logs_Module"))
        vLogsValue      = fOkValue(oRs("Logs_Value"))
        vLogsTimespent  = fOkValue(oRs("Logs_Timespent"))
        vLogsPosted     = fOkValue(oRs("Logs_Posted"))

        If Len(vLogsType) > 0 Then
          If vLogsType = "E" Then
            vTitle = fExamTitle(vLogsModule)
            vScore = vLogsValue
          Else
            vTitle = fModsTitle(vLogsModule)
            '...store the Score and print with next record (timespent)
            If vLogsType = "M" Then
              vScore = vLogsValue
            Else
              vTimeSpent = Cdbl(vLogsTimespent)
            End If
          End If
           
          vTitle = Replace(vTitle, "<b>", "")        
          vTitle = Replace(vTitle, "<B>", "")        
          vTitle = Replace(vTitle, "</b>", "")        
          vTitle = Replace(vTitle, "</B>", "")        
          vTitle = vTitle
        Else
          vTitle = ""
        End If
                
        vPassword = "********"
        If (svMembLevel > vMemb_Level Or vMemb_No = svMembNo) And Instr(vPasswordx, vMemb_Id) = 0 Then
          vPassword = "<a href='User" & fGroup & ".asp?vMembNo=" & vMemb_No & "&vNext=Activity.asp'>" & vMemb_Id & "</a>" & fIf(vMemb_Level = 3, " *", "") & fIf(vMemb_Level = 4, " **", "")
        End If
        
        If vCurList Mod 2 = 0 Then vBg = "bgcolor='#F2F9FD'" Else vBg = ""
    %>
    <tr>
      <td><%=fIf(Len(Trim(vMemb_Criteria)) < 3 Or Trim(vMemb_Criteria) = "0" , "", fCriteria(vMemb_Criteria))%></td>
      <td><div class="ellipsis"><%=fIf(vMemb_Id = vMembIdLast, "", vMemb_FirstName & " " & vMemb_LastName)%></div></td>
      <td><%=fIf(vMemb_Id = vMembIdLast, "", vPassword)%></td>
      <td>
        <%=fIf(vMemb_Active And vMemb_Id <> vMembIdLast, "<img border='0' src='../Images/Icons/CheckMark.gif'>", "")%>
        <%=fIf(Not vMemb_Active And vMemb_Id <> vMembIdLast, "<img border='0' src='../Images/Icons/XMark.gif''>", "")%>
      </td>
      <td><%=fIf(vLogsType = "E", "Exam", vLogsModule)%></td>
      <td><div class="ellipsis"><%=vTitle%></div></td>
      <td><%=vTimeSpent%></td>
      <td><%=vScore%></td>
      <td><%=fFormatDate(vLogsPosted)%></td>
    </tr>
    <%
        vScore = ""
        vTimeSpent = ""
        vCurList = vCurList + 1
        vMembIdLast = vMemb_Id
        If Cint(vCurList) Mod Cint(vMaxList) = 0 Then Exit Do
        oRs.MoveNext
      Loop 
    %>
  </table>

  <div style="text-align: center; margin: 20px;">
    <form method="POST" action="Activity_O.asp">
      <input type="button" onclick="location.href = 'Activity.asp?vStrDate=<%=vStrDate%>&vCurList=<%=vCurList%>&vActive=<%=vActive%>&vMods=<%=vMods%>&vModsOnly=<%=vModsOnly%>&vFind=<%=vFind%>&vFindId=<%=vFindId%>&vFindFirstName=<%=vFindFirstName%>&vFindLastName=<%=vFindLastName%>&vFindEmail=<%=vFindEmail%>&vFindCriteria=<%=vFindCriteria%>&vFormat=<%=vFormat%>'" value="Restart" name="bReturn" id="bReturn" id="bReturn" class="button085">
      <% If Cint(vCurList) > 0 And Cint(vCurList) Mod vMaxList = 0 Then '...If next group, get next starting value %>
      <%=f10%>
      <input type="hidden" name="vStrDate" value="<%=vStrDate%>">
      <input type="hidden" name="vCurList" value="<%=vCurList%>">
      <input type="hidden" name="vMods" value="<%=vMods%>">
      <input type="hidden" name="vModsOnly" value="<%=vModsOnly%>">
      <input type="hidden" name="vFind" value="<%=vFind%>">
      <input type="hidden" name="vActive" value="<%=vActive%>">
      <input type="hidden" name="vFindId" value="<%=vFindId%>">
      <input type="hidden" name="vFindFirstName" value="<%=vFindFirstName%>">
      <input type="hidden" name="vFindLastName" value="<%=vFindLastName%>">
      <input type="hidden" name="vFindEmail" value="<%=vFindEmail%>">
      <input type="hidden" name="vFindMemo" value="<%=vFindMemo%>">
      <input type="hidden" name="vFindCriteria" value="<%=vFindCriteria%>">
      <input type="hidden" name="vFormat" value="<%=vFormat%>">
      <input type="submit" name="bNext" value="<%=bNext%>" class="button085">
      <% End If %>
    </form>
  </div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


