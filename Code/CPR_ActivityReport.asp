<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  '...get form info
  Dim vUsers, vStrDate, vEndDate

  vUsers      = Request("vUsers")

  '...select between start/end dates
  If vUsers   = "" Then
    vStrDate  = fFormatDate(Request("vStrDate"))
    If IsDate(vStrDate) Then 
      vStrDate = Cdate(vStrDate)
    Else
      vUsers = "y"
    End If
    vEndDate = fFormatDate(Request("vEndDate"))
    If IsDate(vEndDate) Then 
      vEndDate = Cdate(vEndDate)
    Else
      vUsers = "y"
    End If
    If vUsers = "" Then
      If vEndDate < vStrDate Then
        vUsers = "y"
      End If
    End If
  End If

  Dim vCnt(3, 2)
  For i = 0 to 3 : vCnt (i, 0) = 0 : vCnt (i, 1) = 0 : vCnt (i, 2) = 0 : Next

  vSql = "SELECT * FROM Memb WITH (nolock) WHERE (Memb_AcctId = '" & svCustAcctId & "')"
 

  sOpenDb
  Set oRs = oDb.Execute(vSql)

  Do While Not oRs.Eof
    sReadMemb

    '...just count users/facilitators/managers
    i = vMemb_Level - 1 
    If i > 0 And i < 4 Then

      '...determine if member is active by two methods...
      If vUsers = "y" Then
        j = 0 : If vMemb_Active Then j = 1
      Else
        '...inactive
        If vMemb_LastVisit >= vStrDate And vMemb_LastVisit <= vEndDate And Not vMemb_Active Then
          j = 0
        '...active 
        ElseIf vMemb_FirstVisit >= vStrDate And vMemb_FirstVisit <= vEndDate And vMemb_Active Then
          j = 1
        '...other (ie joined after the specified dates)
        Else
          j = 2        
        End If
      End If
      
      vCnt(i, j) = vCnt(i, j) + 1
      vCnt(0, j) = vCnt(0, j) + 1 '...totals (level = 0)    
    End If       
    oRs.MoveNext
  Loop

  sCloseDb
%>


<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

  <% Server.Execute vShellHi %>
  <form method="POST" action="CPR_ActivityReport.asp">
    <table border="0" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="0" cellspacing="0">
      <tr>
        <td><h1 align="center">The Learners Onsite Report </h1><h2>This displays who was Active and Inactive during a particular period over the past 12 months or overall.&nbsp; After you have made your selections, click <b>Go</b>.<br><br>Note: If you choose to count all learners, then learners are considered Inactive if they are currently flagged as inactive, otherwise they are considered Active.&nbsp; However: If you select start and end dates, then learners are considered Inactive if their last site access date was between the dates selected and they are currently flagged as Inactive.&nbsp; Learners are considered Active if their first site access date was between the selected dates and they are currently flagged as Active.&nbsp; Learners that were neither active nor inactive during the time selected, but are currently on file, will be considered as Other.</h2></td>
      </tr>
      <tr>
        <td>
        <div align="center">
          <center>
          <table border="0" cellpadding="5" style="border-collapse: collapse" bordercolor="#DDEEF9" width="80%">
            <% If svMembLevel = 5 Then %>
            <% Else %> <input type="hidden" name="vAccounts" value="<%=vAccounts%>"><% End If %>
            <tr>
              <th align="right" valign="top" nowrap>Count ALL Active/Inactive Learners :</th>
              <td valign="top"><input type="checkbox" name="vUsers" value="y" <%=fcheck("y", vusers)%>> or, un-tick and...</td>
            </tr>
            <tr>
              <th align="right" valign="top" nowrap>Count Learners that were&nbsp;&nbsp;&nbsp; <br>&nbsp;Active/Inactive between Start Date :</th>
              <td valign="top"><input type="text" name="vStrDate" size="20" value="<%=fFormatSqlDate(vStrDate)%>"> <br>ie Jan 1, <%=Year(Now)%></td>
            </tr>
            <tr>
              <th align="right" valign="top" nowrap>and End Date :</th>
              <td valign="top"><input type="text" name="vEndDate" size="20" value="<%=fFormatSqlDate(vEndDate)%>"> <br>ie Dec 31, <%=Year(Now)%>.&nbsp; Note: If &quot;ALL&quot; learners is not un-ticked or if either date is not selected, or End Date is before Start Date then count will be for &quot;ALL&quot; current Learners.</td>
            </tr>
            <tr>
              <td align="right" valign="top">&nbsp;</td>
              <td valign="top" align="center"><p><input type="submit" value="Go" name="bGo" class="button"></p><p>&nbsp;</p></td>
            </tr>
          </table>
          </center></div>
        </td>
      </tr>
      <tr>
        <td>&nbsp;<div align="center">
          <center>
          <table border="1" style="border-collapse: collapse" bordercolor="#DDEEF9" id="AutoNumber2" cellspacing="5" cellpadding="7">
            <tr>
              <th width="20%" nowrap>&nbsp;</th>
              <th align="right" width="20%" nowrap>
              <!--webbot bot='PurpleText' PREVIEW='Active'--><%=fPhra(000063)%></th>
              <th align="right" width="20%" nowrap>
              <!--webbot bot='PurpleText' PREVIEW='Inactive'--><%=fPhra(000154)%></th>
              <th align="right" width="20%" nowrap>
              <!--webbot bot='PurpleText' PREVIEW='Other'--><%=fPhra(000209)%></th>
              <th align="right" width="20%" nowrap>
              <!--webbot bot='PurpleText' PREVIEW='Total'--><%=fPhra(000020)%></th>
            </tr>
            <tr>
              <th width="20%" align="right" nowrap>Learners </th>
              <td align="right" width="20%"><%=vCnt(1, 1)%> </td>
              <td align="right" width="20%"><%=vCnt(1, 0)%> </td>
              <td align="right" width="20%"><%=vCnt(1, 2)%> </td>
              <td align="right" width="20%"><%=vCnt(1, 0) + vCnt(1, 1)+ vCnt(1, 2)%> </td>
            </tr>
            <tr>
              <th width="20%" align="right" nowrap>Facilitators </th>
              <td align="right" width="20%"><%=vCnt(2, 1)%> </td>
              <td align="right" width="20%"><%=vCnt(2, 0)%> </td>
              <td align="right" width="20%"><%=vCnt(2, 2)%> </td>
              <td align="right" width="20%"><%=vCnt(2, 0) + vCnt(2, 1) + vCnt(2, 2)%> </td>
            </tr>
            <tr>
              <th width="20%" align="right" nowrap>Managers </th>
              <td align="right" width="20%"><%=vCnt(3, 1)%> </td>
              <td align="right" width="20%"><%=vCnt(3, 0)%> </td>
              <td align="right" width="20%"><%=vCnt(3, 2)%> </td>
              <td align="right" width="20%"><%=vCnt(3, 0) + vCnt(3, 1) + vCnt(3, 2)%> </td>
            </tr>
            <tr>
              <th width="20%" align="right" nowrap>Total </th>
              <td align="right" width="20%"><%=vCnt(0, 1)%></td>
              <td align="right" width="20%"><%=vCnt(0, 0)%></td>
              <td align="right" width="20%"><%=vCnt(0, 2)%></td>
              <td align="right" width="20%"><%=vCnt(0, 0) + vCnt(0, 1) + vCnt(0, 2)%> </td>
            </tr>
          </table>
          </center></div>
        </td>
      </tr>
    </table>
  </form>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>



