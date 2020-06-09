<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Caln.asp"-->

<%
  Dim vM, vMM, vY, aD(56), vDate, vWeekDay, vDesc

  Dim vTskH_Id, vTskH_No, vBg
  vTskH_Id = Request("vTskH_Id")
  vTskH_No = Request("vTskH_No")
  If fNoValue(vTskH_No) Then vTskH_No = 0

  vCaln_Date = Request("vCaln_Date")  

  
  If Request.Form("vM").Count = 1 Then 
    vM = Request.Form("vM")
  ElseIf Not fNoValue(vCaln_Date) Then 
    vM  = Month(vCaln_Date)
  Else
    vM  = Month(now)
  End If

  If Request.Form("vY").Count = 1 Then 
    vY = Request.Form("vY")
  ElseIf Not fNoValue(vCaln_Date) Then 
    vY  = Year(vCaln_Date)
  Else
    vY = Year(now)
  End If

  If Request.Form("bPrev.x").Count = 1 Then   
    If vM > 1 Then
      vM = vM - 1
    Else
      vM = 12
      vY = vY - 1
    End If
  ElseIf Request.Form("bNext.x").Count = 1 Then   
    If vM < 12 Then
      vM = vM + 1
    Else
      vM = 1
      vY = vY + 1
    End If
  End If
  
  vDate = CDate(MonthName(vM) & " 1, " & vY)
  vWeekDay = WeekDay(vDate)
  
  k = -1
  For i = vWeekDay -1 to 56
    k = k + 1
    j = vDate + k
    If Cint(Month(j)) <> Cint(vM) Then Exit For
    aD(i) = MonthName(Month(j), 1) & " " & Day(j)
  Next

%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" id="AutoNumber1">
    <tr>
      <td width="100%">
      <h1><b><font face="Verdana" size="1">C</font></b>alendar</h1>
      <h2><br>Chose the month, if any calendar icons appear, use the &quot;mouse over&quot; to review the activities on that day.&nbsp; To add or edit items to that day, click on the date.&nbsp; Today&#39;s date has a yellow background.<br>&nbsp;</h2>
      </td>
    </tr>
    <tr>
      <td width="100%">
      <form name="fForm" action="CalnList.asp" method="POST">
        <center>
        <table cellspacing="0" cellpadding="0">
          <tr>
            <td rowspan="2"><select name="vM" width="97" size="1">
            <option <%=fselect("1",  vm)%> value="1">January</option>
            <option <%=fselect("2",  vm)%> value="2">February</option>
            <option <%=fselect("3",  vm)%> value="3">March</option>
            <option <%=fselect("4",  vm)%> value="4">April</option>
            <option <%=fselect("5",  vm)%> value="5">May</option>
            <option <%=fselect("6",  vm)%> value="6">June</option>
            <option <%=fselect("7",  vm)%> value="7">July</option>
            <option <%=fselect("8",  vm)%> value="8">August</option>
            <option <%=fselect("9",  vm)%> value="9">September</option>
            <option <%=fselect("10", vm)%> value="10">October</option>
            <option <%=fselect("11", vm)%> value="11">November</option>
            <option <%=fselect("12", vm)%> value="12">December</option>
            </select> <select name="vY" width="97" size="1">
            <option <%=fselect(2000, vy)%>>2000</option>
            <option <%=fselect(2001, vy)%>>2001</option>
            <option <%=fselect(2002, vy)%>>2002</option>
            <option <%=fselect(2003, vy)%>>2003</option>
            <option <%=fselect(2004, vy)%>>2004</option>
            <option <%=fselect(2005, vy)%>>2005</option>
            <option <%=fselect(2006, vy)%>>2006</option>
            <option <%=fselect(2007, vy)%>>2007</option>
            <option <%=fselect(2008, vy)%>>2008</option>
            <option <%=fselect(2009, vy)%>>2009</option>
            <option <%=fselect(2010, vy)%>>2010</option>
            </select> <input border="0" src="../Images/Buttons/Go_<%=svLang%>.gif" name="bGo" type="image">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
            <td valign="middle" align="center"><input border="0" src="../Images/Common/back.gif" name="bPrev" type="image">&nbsp; <input border="0" src="../Images/Common/next.gif" name="bNext" type="image"> </td>
          </tr>
          <tr>
            <td valign="middle" align="center"><font face="Verdana" size="1"><b>Month</b></font></td>
          </tr>
        </table>
        </center><input type="hidden" name="vTskH_Id" value="<%=vTskH_Id%>"><input type="hidden" name="vTskH_No" value="<%=vTskH_No%>">
      </form>
      </td>
    </tr>
  </table>
  <table border="1" width="100%" bordercolor="#DDEEF9" style="border-collapse: collapse" cellpadding="0" cellspacing="0">
    <tr>
      <th><center>Sun</center></th>
      <th><center>Mon</center></th>
      <th><center>Tue</center></th>
      <th><center>Wed</center></th>
      <th><center>Thu</center></th>
      <th><center>Fri</center></th>
      <th><center>Sat</center></th>
    </tr>
    <% For i = 1 to 7 '...go thu the max 7 week span in a month%> <tr>
      <%
      For j = 1 to 7 '...go thru sun/sat
        k = (i-1)*7 + j-1
        If fNoValue(aD(k)) Then
    %> <td valign="top"><br></td>
      <%    
        Else
          '...see if any activities on this day
          sGetCaln Cdate(aD(k) & ", " & vY), vTskH_No
          '...if today then highlight cell with yellow
          vBg = "" : If Cdate(date) = Cdate(aD(k) & ", " & vY) Then vBg=" bgcolor='#FFFF00'"
    %> <td align="center" valign="top" <%=vbg%>><a <%=fStatX%> href="CalnEdit.asp?vCaln_Date=<%=Cdate(aD(k) & ", " & vY) & "&vTskH_Id=" & vTskH_Id & "&vTskH_No=" & vTskH_No%>"><%=aD(k)%></a><br><% If Not vCaln_Eof And Len(vCaln_Details) > 1 Then %> <img border="0" src="../Images/Icons/calendar.gif" alt="<%=vCaln_Details%>"> <% Else %>&nbsp;&nbsp; <% End If %> </td>
      <% 
        End If
      Next
    %> </tr>
    <% 
      If fNoValue(aD(k + 1)) Then Exit For
    Next
  %> <tr>
      <td colspan="7" align="center"><br><a <%=fStatX%> href="MyWorld.asp?vTskH_Id=<%=vTskH_Id%>"><img border="0" src="../Images/Icons/World.gif"></a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a <%=fStatX%> href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a><br>&nbsp;</td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

