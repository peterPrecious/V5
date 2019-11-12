<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<% 
  Dim vCookie, vKey, vUrl, vDetails, vScore1, vScore2, vLevel, vStrDate, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindMemo, vFindCriteria, vFormat, vReport, vParmNo, vSaveCriteria
  vCookie   = svCustAcctId & "_LogReport5"

  '...if we did NOT arrive here from THIS form, check for a previously save cookie or a current cookie if we are returning/restarting from the report(s)
  If Request("vForm").Count = 0 Then

    vDetails       = fDefault(Request.Cookies(vCookie)("vDetails"), "n") 
    vReport        = fIf(vDetails = "s", "S", "A")
    vLevel         = fDefault(Request.Cookies(vCookie)("vLevel"), "2") 
    vScore1        = fDefault(Request.Cookies(vCookie)("vScore1"), "GE")
    vScore2        = fDefault(Request.Cookies(vCookie)("vScore2"), 0)
    vStrDate       = fDefault(Request.Cookies(vCookie)("vStrDate"), fFormatSqlDate(DateAdd ("d", -90, Now)))
    vFind          = fDefault(Request.Cookies(vCookie)("vFind"), "S")
    vFindId        = fUnQuote(Request.Cookies(vCookie)("vFindId"))
    vFindFirstName = fUnQuote(Request.Cookies(vCookie)("vFindFirstName"))
    vFindLastName  = fUnQuote(Request.Cookies(vCookie)("vFindLastName"))
    vFindEmail     = fNoQuote(Request.Cookies(vCookie)("vFindEmail"))
    vFindMemo      = fUnQuote(Request.Cookies(vCookie)("vFindMemo"))
    vFindCriteria  = fDefault(Request.Cookies(vCookie)("vFindCriteria"), "0")
    vParmNo        = fDefault(Request.Cookies(vCookie)("vParmNo"), Request("vParmNo")) '...this value can come from corporate links
    vFormat        = fIf(Request.Cookies(vCookie)("bFormat") = "Online", "O", "X")
    vSaveCriteria  = fDefault(Request.Cookies(vCookie)("vSaveCriteria"), "n")

  '...else assume we arrived her from THIS form...
  Else

    vDetails       = Request.Form("vDetails")
    vReport        = fIf(vDetails = "s", "S", "A")
    vLevel         = fDefault(Request.Form("vLevel"), "2")
    vScore1        = Request.Form("vScore1")
    vScore2        = Request.Form("vScore2")
    vStrDate       = Request.Form("vStrDate")
    vFind          = Request.Form("vFind")
    vFindId        = Ucase(fUnQuote(Request.Form("vFindId")))
    vFindFirstName = fUnQuote(Request.Form("vFindFirstName"))
    vFindLastName  = fUnQuote(Request.Form("vFindLastName"))
    vFindEmail     = fNoQuote(Request.Form("vFindEmail"))
    vFindMemo      = fUnQuote(Request.Form("vFindMemo"))
    vFindCriteria  = Replace(Request.Form("vFindCriteria"), ",", "")
    vParmNo        = Request.Form("vParmNo")
    vFormat        = fIf(Request.Form("bFormat") = "Online", "O", "X")
    vSaveCriteria  = Request.Form("vSaveCriteria")
    
    '...wipe out the cookie
    If Response.Cookies(vCookie).HasKeys Then
      For Each vKey in Response.Cookies(vCookie)
        Response.Cookies(vCookie)(vKey) = ""
      Next
    Else
      Response.Cookies(vCookie) = ""
    End If
        
    '...save cookie for 30 days?
    If vSaveCriteria = "y" Then
      Response.Cookies(vCookie).Expires = DateAdd ("d", 30, Now)
    End If

    '...save the above in a cookie for next reports and Posterity (if chosen)
    Response.Cookies(vCookie)("vDetails")       = vDetails
    Response.Cookies(vCookie)("vReport")        = vReport
    Response.Cookies(vCookie)("vLevel")         = vLevel
    Response.Cookies(vCookie)("vScore1")        = vScore1
    Response.Cookies(vCookie)("vScore2")        = vScore2
    Response.Cookies(vCookie)("vCurList")       = 0
    Response.Cookies(vCookie)("vStrDate")       = vStrDate
    Response.Cookies(vCookie)("vFind")          = vFind
    Response.Cookies(vCookie)("vFindId")        = vFindId
    Response.Cookies(vCookie)("vFindFirstName") = vFindFirstName
    Response.Cookies(vCookie)("vFindLastName")  = vFindLastName
    Response.Cookies(vCookie)("vFindEmail")     = vFindEmail
    Response.Cookies(vCookie)("vFindMemo")      = vFindMemo
    Response.Cookies(vCookie)("vFindCriteria")  = vFindCriteria
    Response.Cookies(vCookie)("vParmNo")        = vParmNo
    Response.Cookies(vCookie)("vFormat")        = vFormat
    Response.Cookies(vCookie)("vSaveCriteria")  = vSaveCriteria

    '...launch the appropriate report
    Session("soRs") = ""  
    vUrl = "LogReport5" & vReport & vFormat & ".asp"
    Response.Redirect vUrl

    '...or debug cookies (comment out previous line)
    For Each vKey in Request.Cookies(vCookie)
      Response.Write vKey & " = " & Request.Cookies(vCookie)(vKey) & "<br>"
    Next

  End If
  
%>

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td valign="top">
      <h1 align="center">
      <!--[[-->Assessment | Survey Report<!--]]--></h1>
      <h2><!--[[-->This will produce a report of assessment scores or survey results&nbsp; which you can either display Online or as an Excel spreadsheet which can be saved on your system.<!--]]-->&nbsp; <!--[[-->If you choose to display Assessments - Highest score achieved you will get the highest score and the last time the learner attempted that assessment. The date may not correspond to the date that the learner got the highest score. For example, a learner may have taken a self assessment on a Tuesday and got 85%. Then they tried to improve their score on Wednesday but only got 80%. This report will show the highest score (85%) and the last date the assessment was taken (Wednesday).<!--]]-->&nbsp;      <!--[[-->The assessment report allows you to generate on online certificate - regardless of the passing mark.<!--]]-->&nbsp;      <!--[[-->If you select Online Surveys the results are displayed in one long string.&nbsp; These are best analyzed in the Excel format which inserts each response into a spreadsheet cell.<!--]]--></h2>
      </td>
    </tr>
    <tr>
      <td valign="top">
      <form method="POST" action="LogReport5.asp">
        <table border="1" width="100%" cellspacing="0" cellpadding="2" bordercolor="#DDEEF9" style="border-collapse: collapse">
          <tr>
            <th align="right" nowrap valign="top">
            <!--[[-->Select<!--]]--> :</th>
            <td>
              <input type="radio" value="y" name="vDetails" <%=fcheck("y", vDetails)%>><!--[[-->Assessments<!--]]--> - <!--[[-->All attempts<!--]]-->&nbsp;&nbsp; <br>
              <input type="radio" value="n" name="vDetails" <%=fcheck("n", vDetails)%>><!--[[-->Assessments<!--]]--> - <!--[[-->Highest score achieved<!--]]--><br>
              <input type="radio" value="s" name="vDetails" <%=fcheck("s", vDetails)%>><!--[[-->Surveys<!--]]--></td>
          </tr>
          <tr>
            <th align="right" nowrap valign="top" height="29">
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
            <!--[[-->taken during last<!--]]--> :</th>
            <td height="29">
              <select size="1" name="vStrDate"><%=vOption%></select></td>
          </tr>
          <%    
            vOption = ""
            vSelected = ""
            For i = 0 To 100 
              vSelected = fIf(Cint(vScore2) = i, " selected", "")
              vOption = vOption & "<option value='" & i & "'" & vSelected & ">" & i & "%</option>" & vbCrLf 
            Next
          %>
          <tr>
            <th align="right" width="50%" valign="top">showing scores that are :</th>
            <td width="50%">
              <table border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><input type="radio" name="vScore1" value="GE" <%=fCheck(vScore1, "GE")%>>&gt;=</td>
                  <td rowspan="2">&nbsp;&nbsp;&nbsp;&nbsp;<select size="1" name="vScore2"><%=vOption%></select></td>
                  <td rowspan="2" valign="bottom">&nbsp;&nbsp; (Does not apply to Surveys)</td>
                </tr>
                <tr>
                  <td><input type="radio" name="vScore1" value="LE" <%=fCheck(vScore1, "LE")%>>&lt;=</td>
                </tr>
              </table>
            </td>
          </tr>
          <tr>
            <th align="right" width="50%" valign="top">from : </th>
            <td width="50%">
              <input type="checkbox" name="vLevel" value="2" <%=fcheck("2", vLevel)%>>Learners<br>
              <input type="checkbox" name="vLevel" value="3" <%=fcheck("3", vLevel)%>>Facilitators<br>
              <% If svMembLevel > 3 Then %>
              <input type="checkbox" name="vLevel" value="4" <%=fcheck("4", vLevel)%>>Managers<br>
              <% End If %>
              <% If svMembLevel = 5 Then %>
              <input type="checkbox" name="vLevel" value="5" <%=fcheck("5", vLevel)%>>Administrators
              <% End If %>
              (if you leave empty, Learners will be selected)</td>
          </tr>
          <tr>
            <th align="right" width="50%" valign="top">that :</th>
            <td width="50%">
              <input type="radio" name="vFind" value="S" <%=fcheck("s", vfind)%>><!--[[-->start with<!--]]--><br> 
              <input type="radio" name="vFind" value="C" <%=fcheck("c", vfind)%>><!--[[-->contain<!--]]--></td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">&nbsp;<%=fIf(svCustPwd, "<!--{{-->Learner Id<!--}}-->", "<!--{{-->Password<!--}}-->")%> : </td>
            <td width="50%"><input type="text" name="vFindId" size="29" value="<%=vFindId%>"></td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">
            <!--[[-->First Name<!--]]--> : </td>
            <td width="50%"><input type="text" name="vFindFirstName" size="29" value="<%=vFindFirstName%>">&nbsp; </td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">
            <!--[[-->Last Name<!--]]--> :</td>
            <td width="50%"><input type="text" name="vFindLastName" size="29" value="<%=vFindLastName%>"></td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">
            <!--[[-->Email Address<!--]]--> :</td>
            <td width="50%"><input type="text" name="vFindEmail" size="29" value="<%=vFindEmail%>"></td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">Memo :</td>
            <td><input type="text" name="vFindMemo" size="29" value="<%=vFindMemo%>"></td>
          </tr>
          <% 
            i = fCriteriaList (svCustAcctId, "REPT:" & svMembCriteria) '...this is from users.asp (works well)
            If vCriteriaListCnt > 1 Then
              If svMembLevel > 2 Then 
          %>
          <tr>
            <th align="right" width="50%" valign="top"><!--[[-->from Group<!--]]--> :</th>
            <td width="50%">
              <select size="<%=vCriteriaListCnt%>" name="vFindCriteria" multiple><%=i%></select>
            </td>
          </tr>
          <%  
              Else 
          %>
          <input type="hidden" name="vFindCriteria" value="<%=svMembCriteria%>">
          <tr>
            <th align="right" width="50%" height="31"><!--[[-->from Group<!--]]--> :</th>
            <td width="50%" height="31"><%=fCriteria (svMembCriteria)%></td>
          </tr>
          <% 
              End If

            Else 
          %> 
          <input type="hidden" name="vFindCriteria" value="0">
          <% 
            End If 
          %>
          </tr>

          <tr>
            <td colspan="2" align="center">
            	&nbsp;<h2>
              <!--[[-->Then select the report format...<!--]]--></h2><p>
              <input type="submit" value="Online" name="bFormat" class="button"><%=f10%>
              <input type="submit" value="Excel" name="bFormat" class="button">
              <input type="hidden" value="<%=vParmNo%>" name="vParmNo"> 
              </p>
            	<p class="c2">
              <input type="checkbox" name="vSaveCriteria" value="y" <%=fCheck("y", vSaveCriteria)%>><!--[[-->Save the above selection criteria <br>on my computer for 30 days.<!--]]--></p>
            </td>
          </tr>
        </table>
        <input type="hidden" name="vForm" value="y">
      </form>
      </td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>