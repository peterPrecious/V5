<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<% 
  Dim vCookie, vKey, vUrl, vDetails, vScore1, vScore2, vLevel, vStrDate, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindCriteria, vFormat, vReport, vParmNo, vSaveCriteria
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
    vFindCriteria  = Replace(Request.Form("vFindCriteria"), ",", "")
    vParmNo        = Request.Form("vParmNo")
    vFormat        = fIf(Request.Form("bFormat") = "Online", "O", "_x")
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
  <meta charset="UTF-8">
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
      <!--webbot bot='PurpleText' PREVIEW='Assessment | Survey Report'--><%=fPhra(000476)%></h1>
      <h2><!--webbot bot='PurpleText' PREVIEW='This will produce a report of assessment scores or survey results&nbsp; which you can either display Online or as an Excel spreadsheet which can be saved on your system.'--><%=fPhra(000477)%>&nbsp; <!--webbot bot='PurpleText' PREVIEW='If you choose to display Assessments - Highest score achieved you will get the highest score and the last time the learner attempted that assessment. The date may not correspond to the date that the learner got the highest score. For example, a learner may have taken a self assessment on a Tuesday and got 85%. Then they tried to improve their score on Wednesday but only got 80%. This report will show the highest score (85%) and the last date the assessment was taken (Wednesday).'--><%=fPhra(000478)%>&nbsp;      <!--webbot bot='PurpleText' PREVIEW='The assessment report allows you to generate on online certificate - regardless of the passing mark.'--><%=fPhra(000479)%>&nbsp;      <!--webbot bot='PurpleText' PREVIEW='If you select Online Surveys the results are displayed in one long string.&nbsp; These are best analyzed in the Excel format which inserts each response into a spreadsheet cell.'--><%=fPhra(000480)%></h2>
      </td>
    </tr>
    <tr>
      <td valign="top">
      <form method="POST" action="AssessmentNew.asp">
        <table border="1" width="100%" cellspacing="0" cellpadding="2" bordercolor="#DDEEF9" style="border-collapse: collapse">
          <tr>
            <th align="right" nowrap valign="top">
            <!--webbot bot='PurpleText' PREVIEW='Select'--><%=fPhra(000275)%> :</th>
            <td>
              <input type="radio" value="y" name="vDetails" <%=fcheck("y", vDetails)%>><!--webbot bot='PurpleText' PREVIEW='Assessments'--><%=fPhra(000481)%> - <!--webbot bot='PurpleText' PREVIEW='All attempts'--><%=fPhra(000067)%>&nbsp;&nbsp; <br>
              <input type="radio" value="n" name="vDetails" <%=fcheck("n", vDetails)%>><!--webbot bot='PurpleText' PREVIEW='Assessments'--><%=fPhra(000481)%> - <!--webbot bot='PurpleText' PREVIEW='Highest score achieved'--><%=fPhra(000473)%><br>
              <input type="radio" value="s" name="vDetails" <%=fcheck("s", vDetails)%>><!--webbot bot='PurpleText' PREVIEW='Surveys'--><%=fPhra(000482)%></td>
          </tr>
          <tr>
            <th align="right" nowrap valign="top" height="29">
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
            <!--webbot bot='PurpleText' PREVIEW='taken during last'--><%=fPhra(000250)%> :</th>
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
              <input type="radio" name="vFind" value="S" <%=fcheck("s", vfind)%>><!--webbot bot='PurpleText' PREVIEW='start with'--><%=fPhra(000463)%><br> 
              <input type="radio" name="vFind" value="C" <%=fcheck("c", vfind)%>><!--webbot bot='PurpleText' PREVIEW='contain'--><%=fPhra(000464)%></td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">&nbsp;<%=fIf(svCustPwd, fPhraH(000411), fPhraH(000211))%> : </td>
            <td width="50%"><input type="text" name="vFindId" size="29" value="<%=vFindId%>"></td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">
            <!--webbot bot='PurpleText' PREVIEW='First Name'--><%=fPhra(000156)%> : </td>
            <td width="50%"><input type="text" name="vFindFirstName" size="29" value="<%=vFindFirstName%>">&nbsp; </td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">
            <!--webbot bot='PurpleText' PREVIEW='Last Name'--><%=fPhra(000163)%> :</td>
            <td width="50%"><input type="text" name="vFindLastName" size="29" value="<%=vFindLastName%>"></td>
          </tr>
          <tr>
            <td align="right" width="50%" valign="top">
            <!--webbot bot='PurpleText' PREVIEW='Email Address'--><%=fPhra(000126)%> :</td>
            <td width="50%"><input type="text" name="vFindEmail" size="29" value="<%=vFindEmail%>"></td>
          </tr>
          <% 
            i = fCriteriaList (svCustAcctId, "REPT:" & svMembCriteria) '...this is from users.asp (works well)
            If vCriteriaListCnt > 1 Then
              If svMembLevel > 2 Then 
          %>
          <tr>
            <th align="right" width="50%" valign="top"><!--webbot bot='PurpleText' PREVIEW='from Group'--><%=fPhra(000565)%> :</th>
            <td width="50%">
              <select size="<%=vCriteriaListCnt%>" name="vFindCriteria" multiple><%=i%></select>
            </td>
          </tr>
          <%  
              Else 
          %>
          <input type="hidden" name="vFindCriteria" value="<%=svMembCriteria%>">
          <tr>
            <th align="right" width="50%" height="31"><!--webbot bot='PurpleText' PREVIEW='from Group'--><%=fPhra(000565)%> :</th>
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
              <!--webbot bot='PurpleText' PREVIEW='Then select the report format...'--><%=fPhra(000483)%></h2><p>
              <input type="submit" value="Online" name="bFormat" class="button"><%=f10%>
              <input type="submit" value="Excel" name="bFormat" class="button">
              <input type="hidden" value="<%=vParmNo%>" name="vParmNo"> 
              </p>
            	<p class="c2">
              <input type="checkbox" name="vSaveCriteria" value="y" <%=fCheck("y", vSaveCriteria)%>><!--webbot bot='PurpleText' PREVIEW='Save the above selection criteria <br>on my computer for 30 days.'--><%=fPhra(000833)%></p>
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

