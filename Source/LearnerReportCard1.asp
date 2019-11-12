<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<% 
  Dim vNext, vEdit, vCustId, vActive, vMods, vModsOnly, vFind, vFindId, vFindFailing, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vFindActive, vFindCompleted, vFindBookmarks, vFormat, vStrDate, vEndDate
  Dim vCurList, vMaxList, vMembIdLast, vBg, vTitle, vLearner, vPassword, vMemb_Last, vUrl, vLearnerName

  vCurList       = fDefault(Request("vCurList"), 0)
  vMaxList       = fDefault(Request("vMaxList"), 100)
  vStrDate       = fDefault(Request("vStrDate"), "Jan 1, 2000")
  vEndDate       = fDefault(Request("vEndDate"), fFormatSqlDate(Now())) 
  vNext          = fDefault(Request("vOriginal"), Request("vNext"))  '...in case we are coming back from user.asp and already have a vNext
  vEdit          = fDefault(Request("vEdit"), "User" & fGroup & ".asp")
  vCustId        = fDefault(Request("vCustId"), svCustId)

  vFind          = Request("vFind")
  vFindId        = Request("vFindId")
  vFindFailing   = Request("vFindFailing")
  vFindCompleted = Request("vFindCompleted")
  vFindBookmarks = Request("vFindBookmarks")
  vFindFirstName = Request("vFindFirstName")
  vFindLastName  = Request("vFindLastName")
  vFindEmail     = Request("vFindEmail")
  vFindMemo      = fNoQuote(Request("vFindMemo"))
  vFindCriteria  = Request("vFindCriteria")
  vFindActive    = Request("vFindActive")
  vFormat        = Request("vFormat")
  
  '...for debugging
  vUrl  = "" _
        & "<br>vCurList:       " & vCurList _
        & "<br>vMaxList:       " & vMaxList _
        & "<br>vStrDate:       " & vStrDate _
        & "<br>vEndDate:       " & vEndDate _
        & "<br>vNext:          " & vNext _
        & "<br>vCustId:        " & vCustId _
        & "<br>vFind:          " & vFind _
        & "<br>vFindId:        " & vFindId _
        & "<br>vFindFailing:   " & vFindFailing _
        & "<br>vFindCompleted: " & vFindCompleted _
        & "<br>vFindBookmarks: " & vFindBookmarks _
        & "<br>vFindFirstName: " & vFindFirstName _
        & "<br>vFindLastName:  " & vFindLastName _
        & "<br>vFindEmail:     " & vFindEmail _
        & "<br>vFindMemo:      " & vFindMemo _
        & "<br>vFindCriteria:  " & vFindCriteria _
        & "<br>vFindActive:    " & vFindActive _
        & "<br>vFindCriteria:  " & vFindCriteria _
        & "<br>vFindActive:    " & vFindActive _
        & "<br>vFormat:        " & vFormat
' Response.Write vUrl
  
  sGetCust vCustId
 
%>


<html>

<head>
  <title>LearnerReportCard1</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <style>
    .table .t1 th:nth-child(1) { width:20%; text-align:left; }
    .table .t1 th:nth-child(2) { width:20%; text-align:left; }
    .table .t1 th:nth-child(3) { width:20%; text-align:left; }
    .table .t1 th:nth-child(4) { width:20%; text-align:left;  }
    .table .t1 th:nth-child(5) { width:20%; text-align:center; }
  </style>

</head>

<body>

  <% Server.Execute vShellHi %>

    <% p0 = fFormatDate(vStrDate) : p1 = fFormatDate(vEndDate) %>
    <h1><!--[[-->Learner Report Card<!--]]--></h1>
    <h2><!--[[-->This report, sorted by Last Name, shows all selected Learners who have accessed content between ^0 and ^1.<!--]]--></h2>
    <h3><!--[[-->Click Details for the Learner Report Card.&nbsp; Click on the Learner's Name to access the Learner's Profile.<!--]]--></h3>       
    <br />      

    <table class="table">

      <tr class="t1">
        <th class="rowshade"><!--[[-->Group<!--]]--></th>
        <th class="rowshade"><!--[[-->Learner's Name<!--]]--></th>
        <th class="rowshade"><%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%></th>
        <th class="rowshade"><!--[[-->Memo<!--]]--></th>
        <th class="rowshade">&nbsp;</th>
      </tr>
  
      <% 
      '...Get initial recordset on first pass and store in session variable
      If vCurList = 0 Then 

        vSql = " SELECT DISTINCT Memb.Memb_No, Memb.Memb_Id, Memb.Memb_FirstName, Memb.Memb_LastName, Memb.Memb_Memo, Memb.Memb_Criteria, Memb.Memb_Level" _
             & " FROM "_
             & "   Memb WITH (NOLOCK) LEFT OUTER JOIN "_
             & "   Logs WITH (NOLOCK) ON Memb.Memb_No = Logs.Logs_MembNo " _
             
             & " WHERE Memb.Memb_AcctId = '" & vCust_AcctId & "' "_

             &   fIf(IsDate(vStrDate),                        " AND (CHARINDEX(Logs.Logs_Type, 'TP') > 0)                   ", "" ) _
             &   fIf(IsDate(vStrDate),                        " AND (Logs.Logs_AcctId = '" & vCust_AcctId & "')             ", "" ) _

             &   fIf(vFind = "S" And Len(vFindId) > 0,        " AND (Memb_Id        LIKE '"      & vFindId        & "%')    ", "" ) _
             &   fIf(vFind = "S" And Len(vFindFirstName) > 0, " AND (Memb_FirstName LIKE '"      & fUnQuote(vFindFirstName) & "%')    ", "" ) _
             &   fIf(vFind = "S" And Len(vFindLastName) > 0,  " AND (Memb_LastName  LIKE '"      & fUnQuote(vFindLastName)  & "%')    ", "" ) _
             &   fIf(vFind = "S" And Len(vFindEmail) > 0,     " AND (Memb_Email     LIKE '"      & vFindEmail     & "%')    ", "" ) _
             &   fIf(vFind = "S" And Len(vFindMemo) > 0,      " AND (Memb_Memo      LIKE '"      & vFindMemo      & "%')    ", "" ) _

             &   fIf(vFind = "C" And Len(vFindId) > 0,        " AND (Memb_Id        LIKE '%"     & vFindId        & "%')    ", "" ) _
             &   fIf(vFind = "C" And Len(vFindFirstName) > 0, " AND (Memb_FirstName LIKE '%"     & fUnQuote(vFindFirstName) & "%')    ", "" ) _
             &   fIf(vFind = "C" And Len(vFindLastName) > 0,  " AND (Memb_LastName  LIKE '%"     & fUnQuote(vFindLastName)  & "%')    ", "" ) _
             &   fIf(vFind = "C" And Len(vFindEmail) > 0,     " AND (Memb_Email     LIKE '%"     & vFindEmail     & "%')    ", "" ) _
             &   fIf(vFind = "C" And Len(vFindMemo) > 0,      " AND (Memb_Memo      LIKE '%"     & vFindMemo      & "%')    ", "" ) _

             &   fIf(vFindCriteria <> "0",                    " AND (CHARINDEX(Memb_Criteria, '" & vFindCriteria  & "') > 0)", "" ) _
             &   fIf(vFindCriteria <> "0",                    " AND (Memb_Criteria <> '0')                                  ", "" ) _

             &   fIf(vFindActive = "a",                       " AND (Memb_Active = 1)                                       ", "" ) _
             &   fIf(vFindActive = "i",                       " AND (Memb_Active = 0)                                       ", "" ) _

             &   fIf(IsDate(vStrDate),                        " AND (Logs.Logs_Posted >= '"      & vStrDate & "')           ", "" ) _
             &   fIf(IsDate(vEndDate),                        " AND (Logs.Logs_Posted <= '"      & vEndDate & "')           ", "" ) _

             & " ORDER BY Memb.Memb_Criteria, Memb.Memb_LastName, Memb.Memb_FirstName"
 
'       sDebug
        sOpenDb
        Set oRs = oDB.Execute(vSql)
        Set Session("soRs") = oRs
        vCurList = 1
      Else  
        Set oRs = Session("soRs")
      End If  

      '...read until either eof or end of group
      Do While Not oRs.Eof

        vMemb_No        = oRs("Memb_No")
        vMemb_Id        = oRs("Memb_Id")
        vMemb_FirstName = oRs("Memb_FirstName")
        vMemb_LastName  = oRs("Memb_LastName")
        vMemb_Memo      = oRs("Memb_Memo")
        vMemb_Criteria  = oRs("Memb_Criteria")
        vMemb_Level     = oRs("Memb_Level")

        vLearnerName    = vMemb_FirstName & " " & vMemb_LastName 
        If Trim(vLearnerName) = "" Then vLearnerName = " ...n/a..."

        If (svMembLevel = 5) Or ((svMembLevel > vMemb_Level Or vMemb_No = svMembNo) And InStr(vMemb_Id, vPasswordx) = 0) Then
          '...if we currently have a vNext (ie a return program) then we need to protect this as we
          '   will be sending a new vNext (ie this current page) to the User.asp page.
          '   thus store current vNext in vOriginal so we can reset the vNext when we return here
          vUrl  = "LearnerReportCard1.asp" _
                & "?vMemb_No="       & vMemb_No _
                & "&vStrDate="       & vStrDate _
                & "&vEndDate="       & vEndDate _
                & "&vCurList="       & 0 _
                & "&vOriginal="      & vNext _
                & "&vEdit="          & vEdit _
                & "&vCustId="        & vCustId _
                & "&vFind="          & vFind _
                & "&vFindId="        & vFindId _
                & "&vFindFailing="   & vFindFailing _
                & "&vFindCompleted=" & vFindCompleted _
                & "&vFindBookmarks=" & vFindBookmarks _
                & "&vFindFirstName=" & vFindFirstName _
                & "&vFindLastName="  & vFindLastName _
                & "&vFindEmail="     & vFindEmail _
                & "&vFindMemo="      & vFindMemo _
                & "&vFindCriteria="  & vFindCriteria _
                & "&vFindActive="    & vFindActive _
                & "&vFormat="        & vFormat        
'         Response.Write "<p align='left'>" & Replace(vUrl, "&", "<br>&") & "</p>" 
          vUrl = Server.UrlEncode(vUrl)
          vLearner = "<a href='" & vEdit & "?vMembNo=" & vMemb_No & "&vCustId=" & vCustId & "&vNext=" & vUrl & "'>" & vLearnerName & "</a>" & fIf(vMemb_Level = 3, " *", "") & fIf(vMemb_Level = 4, " **", "") & fIf(vMemb_Level = 5, " ***", "")

          vPassword = fIf(InStr(vMemb_Id, vPasswordx) = 0, vMemb_Id, "********")


        Else
          vLearner  = vLearnerName  & fIf(vMemb_Level = 3, " *", "") & fIf(vMemb_Level = 4, " **", "") & fIf(vMemb_Level = 5, " ***", "")
          vPassword = "********"
        End If  

        If Cint(vCurList) Mod 2 = 0 Then vBg = "bgcolor='#F2F9FD'" Else vBg = ""
        vUrl  = "vMemb_No="        & vMemb_No _
              & "&vStrDate="       & vStrDate _
              & "&vEndDate="       & vEndDate _
              & "&vCurList="       & vCurList _
              & "&vNext="          & vNext _
              & "&vEdit="          & vEdit _
              & "&vCustId="        & vCustId _
              & "&vFind="          & vFind _
              & "&vFindId="        & vFindId _
              & "&vFindFailing="   & vFindFailing _
              & "&vFindCompleted=" & vFindCompleted _
              & "&vFindBookmarks=" & vFindBookmarks _
              & "&vFindFirstName=" & fjUnquote(vFindFirstName) _
              & "&vFindLastName="  & fjUnquote(vFindLastName) _
              & "&vFindEmail="     & vFindEmail _
              & "&vFindMemo="      & vFindMemo _
              & "&vFindCriteria="  & vFindCriteria _
              & "&vFindActive="    & vFindActive _              
              & "&vFormat="        & vFormat         
'       Response.Write "<p align='left'>" & Replace(vUrl, "&", "<br>&") & "</p>" 
'       vUrl = Server.UrlEncode(vUrl)
    %>

      <tr>
        <td><%=fIf(Len(Trim(vMemb_Criteria)) < 3 Or Trim(vMemb_Criteria) = "0" , "", fCriteria(vMemb_Criteria))%></td>
        <td><%=vLearner%></td>
        <td><%=vPassword%></td>
        <td><%=fIf(Instr(vMemb_Memo, "|")>0, vMemb_Memo, "")%></td>
        <td style="text-align:center"><input type="button" onclick="location.href='LearnerReportCard2.asp?<%=vUrl%>'" value="<%=bDetails%>" name="bDetails" class="button100"></td>
      </tr>
      <%
          vMemb_Last = vMemb_No
          vCurList = vCurList + 1
          If Cint(vCurList) Mod Cint(vMaxList) = 0 Then Exit Do
          oRs.MoveNext
        Loop 
      %>

      <% If Cint(vCurList) > 0 And Cint(vCurList) Mod vMaxList = 0 Then '...If next group, get next starting value %>
      <tr>
        <th colspan="5">
          <form method="POST" action="LearnerReportCard1.asp">
            <input type="hidden" name="vNext"           value="<%=vNext%>">
            <input type="hidden" name="vEdit"           value="<%=vEdit%>">
            <input type="hidden" name="vCustId"         value="<%=vCustId%>">
            <input type="hidden" name="vStrDate"        value="<%=vStrDate%>">
            <input type="hidden" name="vEndDate"        value="<%=vEndDate%>">
            <input type="hidden" name="vCurList"        value="<%=vCurList%>">
            <input type="hidden" name="vFind"           value="<%=vFind%>">
            <input type="hidden" name="vFindId"         value="<%=vFindId%>">
            <input type="hidden" name="vFindFailing"    value="<%=vFindFailing%>">
            <input type="hidden" name="vFindCompleted"  value="<%=vFindCompleted%>">
            <input type="hidden" name="vFindBookmarks"  value="<%=vFindBookmarks%>">
            <input type="hidden" name="vFindFirstName"  value="<%=vFindFirstName%>">
            <input type="hidden" name="vFindLastName"   value="<%=vFindLastName%>">
            <input type="hidden" name="vFindEmail"      value="<%=vFindEmail%>">
            <input type="hidden" name="vFindMemo"       value="<%=vFindMemo%>">
            <input type="hidden" name="vFindCriteria"   value="<%=vFindCriteria%>">
            <input type="hidden" name="vFindActive"     value="<%=vFindActive%>">
            <input type="hidden" name="vFormat"         value="<%=vFormat%>">
            <br>
            <input type="submit" name="bNext"           value="<%=bNext%>" class="button100">
          </form>
        </th>
      </tr>
      <% End If %> 

      <% If Cint(vCurList) = 1 Then %>
      <tr>
        <th colspan="5"><!--[[-->No learners have been selected.<!--]]--></th>
      </tr>
      <% End If %> 


    </table>

    <div style="text-align:center; margin:20px;">
      <% If Len(vNext) > 0 Then %>
      <input type="button" onclick="location.href = '<%=vNext%>'" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button100"><%=f10%>
      <% End If %>
      <% 
          vUrl = "LearnerReportCard.asp?" _
              & "vStrDate="         & vStrDate _
              & "&vEndDate="        & vEndDate _
              & "&vCurList="        & vCurList _
              & "&vNext="           & vNext _
              & "&vEdit="           & vEdit _
              & "&vCustId="         & vCustId _
              & "&vFind="           & vFind _
              & "&vFindId="         & vFindId _
              & "&vFindFailing="    & vFindFailing _
              & "&vFindCompleted="  & vFindCompleted _
              & "&vFindBookmarks="  & vFindBookmarks _
              & "&vFindFirstName="  & fjUnquote(vFindFirstName) _
              & "&vFindLastName="   & fjUnquote(vFindLastName) _
              & "&vFindEmail="      & vFindEmail _
              & "&vFindMemo="       & vFindMemo _
              & "&vFindCriteria="   & vFindCriteria _
              & "&vFindActive="     & vFindActive _
              & "&vFormat="         & vFormat         
          'Response.Write "<p align='left'>" & Replace(vUrl, "&", "<br>&") & "</p>"     
      %>
      <input type="button" onclick="location.href = '<%=vUrl%>';" value="<%=bRestart%>" name="bRestart" class="button100">       
    </div>

    <div style="text-align:center; margin:20px;"><%=vCust_Id & "  (" & vCust_Title & ")"%></div>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>