<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<% 
  Dim vAccounts, vInactive, vPrograms, vUrl, vFormat
  Dim vStrDate, vEndDate, vStrDateErr, vEndDateErr

  '...default to previous month when you enter this site
  If Request("vStrDate").Count = 0 And Request("vEndDate").Count = 0 Then
    vStrDateErr = "" : vStrDate = MonthName(Month(Now), True) & " 1, " & Year(Now)
    vEndDateErr = "" : vEndDate = fFormatSqlDate(DateAdd("d", 1, Now))
  Else
    vStrDate  = fFormatSqlDate(Request("vStrDate")) 
    If vStrDate = " " Then 
      vStrDate  = Request("vStrDate") '...put back bad date for display
      vStrDateErr = "Error"
    End If
    vEndDate  = fFormatSqlDate(Request("vEndDate"))
    If vEndDate = " " Then
      vEndDate  = Request("vEndDate")  '...put back bad date for display
      vEndDateErr = "Error"
    End If
    If (Len(vStrDate) > 0 And vStrDateErr = "") And (Len(vEndDate) > 0 And vEndDateErr = "") Then
      If DateDiff("d", vStrDate, vEndDate) < 0 Then
        vEndDateErr = "Error"
      End If
    End If
  End If


  vInactive     = fDefault(Request("vInactive"), "n")
  vAccounts     = fDefault(Replace(Request("vAccounts"), ", ", ""), "ghc")
  vPrograms     = fDefault(Replace(Request("vPrograms"), ", ", " "), "All")
  vFormat       = fDefault(Request("vFormat"), "1")

  '...processing the form?
  If Request.Form.Count > 0 Then
    '...goto online or excel reports
    vUrl   = "ClientActivityReport_" & vFormat  & ".asp" _   
           & "?vAccounts="    & vAccounts  _
           & "&vInactive="    & vInactive  _
           & "&vStrDate="     & fUrlEncode(vStrDate)  _
           & "&vEndDate="     & fUrlEncode(vEndDate)  _
           & "&vPrograms="    & fUrlEncode(vPrograms) _
           & "&vFormat="      & fUrlEncode(vFormat)  
'   Response.Write vUrl 
    Response.Redirect vUrl 
  End If

  '...get all Programs from the big view
  Function fProgOptions
    fProgOptions = ""
    vSql = "SELECT DISTINCT vCustProg_All.Program, vCustProg_All.Title FROM vCustProg_All WHERE (vCustProg_All.Cust LIKE '" & Left(svCustId, 4) & "%') OR (vCustProg_All.Agent = '" & Left(svCustId, 4) & "') "
'   sDebug     
    sOpenDb
    Set oRs = oDb.Execute(vSql)    
    Do While Not oRs.Eof 
      fProgOptions = fProgOptions & "<option value='" & oRs("Program") & "'>" & oRs("Program") & " - " & oRs("Title") & "</option>" & vbCrLf
      oRs.MoveNext
    Loop      
    sCloseDb           
  End Function
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <script Language="JavaScript">
    function Validate(theForm)
    {
      if (theForm.vPrograms.selectedIndex == undefined) 
      {
        return (true);
      }

      if (theForm.vPrograms.selectedIndex < 0)
      {
        alert("Please select one of the \"Program\" options.");
        theForm.vPrograms.focus();
        return (false);
      }
    
      var numSelected = 0;
      var i;
      for (i = 0;  i < theForm.vPrograms.length;  i++)
      {
        if (theForm.vPrograms.options[i].selected)
            numSelected++;
      }
      if (numSelected < 1)
      {
        alert("Please select at least 1 of the \"Program\" options.");
        theForm.vPrograms.focus();
        return (false);
      }
    
      if (numSelected > 50)
      {
        alert("Please select at most 50 of the \"Program\" options.");
        theForm.vPrograms.focus();
        return (false);
      }
      return (true);
    }
    
    function emptyField(vField) {
      fForm(vField).value = "";
    }
    
  </script>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td valign="top">
      <h1 align="center"><br>Client Activity Report</h1>
      <h2 align="center">The Report shows the selected Programs for the selected <b> <%=Left(svCustId, 4)%></b> Sites.</h2>
      </td>
    </tr>
    <tr>
      <td valign="top">
      <form method="POST" action="ClientActivityReport.asp" onsubmit="return Validate(this)" id="fForm">
        <table border="0" width="100%" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
          <tr>
            <th align="right" valign="top" nowrap width="35%">Site type :</th>
            <td width="65%">
              <input type="checkbox" name="vAccounts" value="h" <%=fChecks(vAccounts, "h")%>>Self Service (Channel)<br>
              <input type="checkbox" name="vAccounts" value="g" <%=fChecks(vAccounts, "g")%>>Group (G1 and G2)<br>
              <input type="checkbox" name="vAccounts" value="c" <%=fChecks(vAccounts, "c")%>>Custom (Corporate)
            </td>
          </tr>
          <tr>
            <th align="right" valign="top" width="30%">
            Sites Setup between :</th>
            <td>
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td nowrap>
                  <input type="text" onblur="refillField('vStrDate', 'Jan 1, 2000')" name="vStrDate0" id="vStrDate" size="12" value="<%=vStrDate%>" style="text-align: center" class="c2"> 
                  <a title="<!--webbot bot='PurpleText' PREVIEW='Start at the beginning...'--><%=fPhra(001221)%>" class="debug" onclick="fillField('vStrDate', 'Jan 1, 2000')" href="#">&#937;</a> 
                  <a href="javascript:show_calendar('vStrDate','EN', '<%=Month(Now)-1%>', '<%=Year(Now)%>', 'MONTH DD YYYY');">
                  <img border="0" src="/V5/Images/Icons/Calendar.jpg" align="absbottom"></a> 
                </td>
                <td width="110"><span style="background-color: #FFFF00"><%=vStrDateErr%></span></td>
              </tr>
              <tr>
                <td colspan="2">ex Oct 1, <%=Year(Now - 1)%>
                <!--webbot bot='PurpleText' PREVIEW='(MMM D, YYYY)'--><%=fPhra(000818)%> in English format</td>
              </tr>
            </table>
            </td>
          </tr>
          <tr>
            <th align="right" valign="top" width="30%">
            and :</th>
            <td>
            <table border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td nowrap>
                  <input type="text"                                               onblur="refillField('vEndDate', '<%=fFormatDate(Now)%>')" name="vEndDate0" id="vEndDate" size="12" value="<%=vEndDate%>" style="text-align: center" class="c2"> 
                  <a title="<!--webbot bot='PurpleText' PREVIEW='End at today's date...'--><%=fPhra(000958)%>" class="debug" onclick="fillField('vEndDate', '<%=fFormatDate(DateAdd("d", 1, Now))%>')" href="#">&#937;</a> 
                  <a href="javascript:show_calendar('vEndDate','EN', '<%=Month(Now)-1%>', '<%=Year(Now)%>');">
                  <img border="0" src="/V5/Images/Icons/Calendar.jpg" align="absbottom">
                  </a>
                </td>
                <td width="103">
                  <span style="background-color: #FFFF00"><%=vEndDateErr%></span>
                </td>
              </tr>
              <tr>
                <td colspan="2">ex Dec 31, <%=Year(Now - 1)%>
                <!--webbot bot='PurpleText' PREVIEW='(MMM D, YYYY)'--><%=fPhra(000818)%> in English format</td>
              </tr>
            </table>
            </td>
          </tr>
          <tr>
            <td align="center" valign="top" nowrap colspan="2">&nbsp;</td>
          </tr>
          <tr>
            <td align="center" valign="top" nowrap colspan="2">&nbsp;<p>Select either &quot;All&quot; Programs or one or more individual Programs.<br><br>
              <select name="vPrograms" multiple size="20" style="width: 500" class="c2">
              <option selected value="ALL">All</option>
              <%=fProgOptions%>
              </select><br><br>&nbsp;</td>
          </tr>
          </tr>
          <tr>
            <th align="right" valign="top" nowrap width="35%">Output as : </th>
            <td width="65%">
              <input type="radio" name="vFormat" value="1" <%=fcheck("1", vformat)%> checked>HTML (Online)<br>
              <input type="radio" name="vFormat" value="X" <%=fcheck("X", vformat)%>>Excel (Maximum 2000 rows) <font color="#FF0000">[inactive]</font></td>
          </tr>
          <tr>
            <th height="75" colspan="2"><input type="submit" value="<%=bNext%>" name="bNext" class="button"></th>
          </tr>
        </table>
      </form>
      </td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>

</html>

