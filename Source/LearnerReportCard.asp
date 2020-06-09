<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<% 
  Dim vNext, vEdit, vCustId, vUrl, vFormat, vCurList, vMaxList
  Dim vFind, vFindId, vFindFailing, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vFindActive, vFindCompleted

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

  vNext          = Request("vNext")
  vEdit          = fDefault(Request("vEdit"), "User" & fGroup & ".asp")
  vCustId        = fDefault(Request("vCustId"), svCustId)
  vFind          = fDefault(Request("vFind"), "S")
  vFindId        = fNoQuote(Request("vFindId"))
  vFindFailing   = fDefault(Request("vFindFailing"), "y")
  vFindCompleted = fDefault(Request("vFindCompleted"), "n")
  vFindFirstName = Request("vFindFirstName")
  vFindLastName  = Request("vFindLastName")
  vFindEmail     = fNoQuote(Request("vFindEmail"))
  vFindMemo      = fUnQuote(Request("vFindMemo"))
  vFindCriteria  = fDefault(Replace(Request("vFindCriteria"), ", ", " "), "0")
  vFindActive    = fDefault(Request("vFindActive"), "b")

  If Instr(" " & vFindCriteria & " ", " 0 ") > 0 Then vFindCriteria = "0"
  vFormat        = fDefault(Request("vFormat"), "1")

  vUrl  = "" _
        & "<br>vCurList: "       & vCurList _
        & "<br>vMaxList: "       & vMaxList _
        & "<br>vStrDate: "       & vStrDate _
        & "<br>vEndDate: "       & vEndDate _
        & "<br>vNext: "          & vNext _
        & "<br>vCustId: "        & vCustId _
        & "<br>vFind: "          & vFind _
        & "<br>vFindId: "        & vFindId _
        & "<br>vFindFailing: "   & vFindFailing _
        & "<br>vFindCompleted: " & vFindCompleted _
        & "<br>vFindFirstName: " & vFindFirstName _
        & "<br>vFindLastName: "  & vFindLastName _
        & "<br>vFindEmail: "     & vFindEmail _
        & "<br>vFindMemo: "      & vFindMemo _
        & "<br>vFindCriteria: "  & vFindCriteria _
        & "<br>vFindActive: "    & vFindActive _
        & "<br>vFindCriteria: "  & vFindCriteria _
        & "<br>vFindActive: "    & vFindActive _
        & "<br>vFormat: "        & vFormat
'  Response.Write vUrl


  '...processing the form?
  If Request.Form.Count > 0 Then
    '...goto online or excel reports
    vUrl   = "LearnerReportCard" & vFormat  & ".asp"          _   
           & "?vStrDate="         & vStrDate       _
           & "&vEndDate="         & vEndDate       _
           & "&vNext="            & vNext          _
           & "&vEdit="            & vEdit          _
           & "&vCustId="          & vCustId        _
           & "&vFind="            & vFind          _
           & "&vFindId="          & vFindId        _
           & "&vFindFailing="     & vFindFailing   _
           & "&vFindCompleted="   & vFindCompleted _
           & "&vFindFirstName="   & vFindFirstName _
           & "&vFindLastName="    & vFindLastName  _
           & "&vFindEmail="       & vFindEmail     _
           & "&vFindMemo="        & vFindMemo      _
           & "&vFindCriteria="    & vFindCriteria  _
           & "&vFindActive="      & vFindActive    _
           & "&vFormat="          & vFormat  
'   Response.Write vUrl
    Response.Redirect vUrl 
  End If
  
  sGetCust vCustId

%>
<html>

<head>
  <title>LearnerReportCard</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script src="/V5/Inc/Calendar.js"></script>
  <script>
    function Validate(theForm)
    {
      if (theForm.vFindCriteria.selectedIndex == undefined) 
      {
        return (true);
      }

      if (theForm.vFindCriteria.selectedIndex < 0)
      {
        alert("Please select one of the \"Group\" options.");
        theForm.vFindCriteria.focus();
        return (false);
      }
    
      var numSelected = 0;
      var i;
      for (i = 0;  i < theForm.vFindCriteria.length;  i++)
      {
        if (theForm.vFindCriteria.options[i].selected)
            numSelected++;
      }
      if (numSelected < 1)
      {
        alert("Please select at least 1 of the \"Group\" options.");
        theForm.vFindCriteria.focus();
        return (false);
      }
    
      if (numSelected > 50)
      {
        alert("Please select at most 50 of the \"Group\" options.");
        theForm.vFindCriteria.focus();
        return (false);
      }
      return (true);
    }

  </script>
</head>

<body>

  <% Server.Execute vShellHi %>

  <h1><!--[[-->Learner Report Card<!--]]--></h1>
  <h2><!--[[-->The Learner Report Card shows the detailed learning activities of Learners who meet the selection criteria below.<!--]]--></h2>
  <h3><!--[[-->Clicking on the <span style="color:#FFA500">&#937;</span> icon allows you to set the Start Date to Jan 1, 2000 or the End Date to today.<!--]]--></h3>
  <p>&nbsp;</p>

    <form method="POST" action="LearnerReportCard.asp" id="fForm" name="fForm" onsubmit="return Validate(this)">
      <input type="hidden" name="vNext" value="<%=vNext%>">
      <input type="hidden" name="vEdit" value="<%=vEdit%>">
      <input type="hidden" name="vCustId" value="<%=vCustId%>">
      <table class="table">
        <tr>
          <th colspan="2" style="text-align:center;" class="c3"><!--[[-->Select Learners with activity occurring between...<!--]]--></th>
          <th style="text-align:left;"><!--[[-->Applies to<br />Format...<!--]]--></th>
        </tr>
        <tr>
          <th><!--[[-->Start Date<!--]]--> :</th>
          <td>
          <table>
            <tr>
              <td>
                <input type="text" onblur="refillField('vStrDate', 'Jan 1, 2000')" name="vStrDate" id="vStrDate" size="10" value="<%=vStrDate%>" style="text-align: center"> 
                <a title="<!--[[-->Start at the beginning...<!--]]-->" class="debug" onclick="fillField('vStrDate', 'Jan 1, 2000')" href="#">&#937;</a> 
                <a href="javascript:show_calendar('vStrDate','EN', '<%=Month(Now)-1%>', '<%=Year(Now)%>', 'MONTH DD YYYY');"><img border="0" src="/V5/Images/Icons/Calendar.jpg" style="vertical-align:baseline"></a> 
              </td>
              <td><span style="background-color: #FFFF00"><%=vStrDateErr%></span></td>
            </tr>
            <tr>
              <td colspan="2"><!--[[-->(Mmm D, YYYY) in English format<!--]]--></td>
            </tr>
          </table>
          </td>
          <td>
          <span style="background-color: #00FFFF">1</span>&nbsp;<span style="background-color: #FFFF00">2</span>&nbsp; <span style="background-color: #FFC1C1">3</span></td>
        </tr>
        <tr>
          <th><!--[[-->End Date<!--]]--> :</th>
          <td>
            <table>
              <tr>
                <td><input type="text" onblur="refillField('vEndDate', '<%=fFormatDate(Now)%>')" name="vEndDate" id="vEndDate" size="10" value="<%=vEndDate%>" style="text-align: center"> 
                  <a title="<!--[[-->End at today's date...<!--]]--" class="debug" onclick="fillField('vEndDate',  '<%=fFormatDate(DateAdd("d", 1, Now))%>')" href="#">&#937;</a> 
                  <a href="javascript:show_calendar('vEndDate','EN', '<%=Month(Now)-1%>', '<%=Year(Now)%>');"><img border="0" src="/V5/Images/Icons/Calendar.jpg"  style="vertical-align:baseline"></a> 
                </td>
                <td><span style="background-color: #FFFF00"><%=vEndDateErr%></span> </td>
              </tr>
              <tr>
                <td colspan="2"><!--[[-->(Mmm D, YYYY)<!--]]--> in English format</td>
              </tr>
            </table>
          </td>
          <td>
          <span style="background-color: #00FFFF">1</span>&nbsp;<span style="background-color: #FFFF00">2</span>&nbsp; <span style="background-color: #FFC1C1">3</span></td>
        </tr>
        <tr>
          <th>
          <!--[[-->Showing Failing Scores<!--]]--> :</th>
          <td>
            <input type="radio" name="vFindFailing" value="y" <%=fcheck("y", vfindfailing)%> checked><!--[[-->Yes<!--]]-->&nbsp;&nbsp;&nbsp;&nbsp; 
            <input type="radio" name="vFindFailing" value="n" <%=fcheck("n", vfindfailing)%>><!--[[-->No, leave blank<!--]]-->
          </td>
          <td style="text-align:center"><span style="background-color: #00FFFF">1</span></td>
        </tr>
        <tr>
          <th>
          <!--[[-->For learners that are<!--]]-->: </th>
          <td>
            <input type="radio" name="vFindActive" value="a" <%=fcheck("a", vfindactive)%>><!--[[-->Active<!--]]-->&nbsp; 
            <input type="radio" name="vFindActive" value="i" <%=fcheck("i", vfindactive)%>><!--[[-->Inactive<!--]]-->&nbsp; 
            <input type="radio" name="vFindActive" value="b" <%=fcheck("b", vfindactive)%>><!--[[-->Both Active and Inactive<!--]]--> 
          </td>
          <td>
            <span style="background-color: #00FFFF">1</span>&nbsp; 
            <span style="background-color: #FFFF00">2</span>&nbsp; 
            <span style="background-color: #FFC1C1">3</span></td>
        </tr>
        <tr>
          <th><!--[[-->Filter for learners that<!--]]--> :</th>
          <td>
            <input type="radio" name="vFind" value="S" <%=fcheck("s", vfind)%>><!--[[-->start with<!--]]-->&nbsp; or 
            <input type="radio" name="vFind" value="C" <%=fcheck("c", vfind)%>><!--[[-->contain<!--]]-->
          </td>
          <td>
            <span style="background-color: #00FFFF">1</span>&nbsp; <span style="background-color: #FFFF00">2</span></td>
        </tr>
        <tr>
          <th><%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%> :</th>
          <td><input type="text" name="vFindId" size="29" value="<%=vFindId%>"></td>
          <td><span style="background-color: #00FFFF">1</span>&nbsp; <span style="background-color: #FFFF00">2</span></td>
        </tr>
        <tr>
          <th><!--[[-->First Name<!--]]--> :</th>
          <td><input type="text" name="vFindFirstName" size="29" value="<%=vFindFirstName%>"></td>
          <td><span style="background-color: #00FFFF">1</span>&nbsp; <span style="background-color: #FFFF00">2</span></td>
        </tr>
        <tr>
          <th><!--[[-->Last Name<!--]]--> :</th>
          <td><input type="text" name="vFindLastName" size="29" value="<%=vFindLastName%>"></td>
          <td><span style="background-color: #00FFFF">1</span>&nbsp; <span style="background-color: #FFFF00">2</span></td>
        </tr>
        <tr>
          <th><!--[[-->Email Address<!--]]--> :</th>
          <td><input type="text" name="vFindEmail" size="29" value="<%=vFindEmail%>"></td>
          <td><span style="background-color: #00FFFF">1</span>&nbsp; <span style="background-color: #FFFF00">2</span></td>
        </tr>
        <tr>
          <th><!--[[-->Memo<!--]]--> :</th>
          <td><input type="text" name="vFindMemo" size="29" value="<%=vFindMemo%>"></td>
          <td><span style="background-color: #00FFFF">1</span>&nbsp; <span style="background-color: #FFFF00">2</span></td>
        </tr>
        <% 
          i = fCriteriaList (vCust_AcctId, "REPT:" & svMembCriteria)
          If vCriteriaListCnt > 1 Then
        %>
        <tr>
          <th>
          <!--[[-->from Group<!--]]--> :</th>
          <td><select size="<%=vCriteriaListCnt%>" name="vFindCriteria" multiple><%=i%></select></td>
          <td><span style="background-color: #00FFFF">1</span>&nbsp; <span style="background-color: #FFFF00">2</span></td>
        </tr>
        <%  
            Else 
        %> <input type="hidden" name="vFindCriteria" value="<%=svMembCriteria%>">
        <tr>
          <th>
          <!--[[-->from Group<!--]]--> :</th>
          <td><%=fCriteria (svMembCriteria)%></td>
          <td><span style="background-color: #00FFFF">1</span>&nbsp; <span style="background-color: #FFFF00">2</span></td>
        </tr>
        <% 
          End If 
        %>
        <tr>
          <th>Format : </th>
          <td colspan="2">
            <span style="background-color: #00FFFF">1.</span><input type="radio" name="vFormat" value="1" <%=fcheck("1", vformat)%> checked>Online<br>
            <span style="background-color: #FFFF00">2.</span><input type="radio" name="vFormat" value="_x" <%=fcheck("_x", vformat)%>>Excel Complete&nbsp; (maximum 5000 rows) <font color="#FF0000">&nbsp;</font>&nbsp; <br>
            <span style="background-color: #FFC1C1">3.</span><input type="radio" name="vFormat" value="_xx" <%=fcheck("_xx", vformat)%>>Excel Scores (maximum 50,000 rows)
          </td>
        </tr>

        </table>

        <div style="text-align:center;">

          <div class="c6" style="font-weight:normal; margin:20px; padding:10px; border:1px solid red;">
            Note: The Format options above right show which selection criteria work with which format. <b>Excel Complete</b> essentially returns the same details as the Online version, to a maximum of 5000 rows. <b>Excel Scores</b> returns <b>scores only</b> for all active and/or inactive learners in the selected date range to a maximum of 50,000 rows. It will disregard any selection criteria from the choices above other than date range and active/inactive status. 
          </div>
          <% If Len(vNext) > 0 Then %> 
          <input type="button" onclick="location.href = '<%=vNext%>'" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button100"><%=f10%> 
          <% End If %> 
          <input type="submit" value="<%=bNext%>" name="bNext" class="button100">
          <br><br><%=vCust_Id & "  (" & vCust_Title & ")"%>
        </div>



    </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>