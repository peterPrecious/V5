<!--#include virtual = "V5\Inc\Setup.asp"-->
<!--#include virtual = "V5\Inc\Initialize.asp"-->
<!--#include virtual = "V5\Inc\Db_Cust.asp"-->
<!--#include virtual = "V5\Inc\Db_Phra.asp"-->
<!--#include virtual = "V5\Inc\Db_Memb.asp"-->
<!--#include virtual = "V5\Inc\Db_Crit.asp"-->

<% 
  Dim vUrl, vCurList, vStrDate, vFind, vFindId, vFindFirstName, vFindLastName, vFindEmail, vFindCriteria, vFormat
  Dim vOption, vSelected, vCF, aCF, vPrev, vDesc

  vCurList        = fDefault(Request("vCurList"), 0)
  vStrDate        = fDefault(Request("vStrDate"), fFormatSqlDate(DateAdd ("d", -999, Now)))
  vFind           = fDefault(Request("vFind"), "S")
  vFindId         = fUnQuote(Request("vFindId"))
  vFindFirstName  = fUnQuote(Request("vFindFirstName"))
  vFindLastName   = fUnQuote(Request("vFindLastName"))
  vFindCriteria   = fDefault(Request("vFindCriteria"), "0") 
  vFindEmail      = fNoQuote(Request("vFindEmail"))
  vFormat         = fIf(Request("bExcel").Count = 1, "X", "O")

  '...processing the form?
  If Request("vForm").Count = 1 Then
    Session("soRs") = "" 
    '...goto online or excel reports
    vUrl = "ReportCard_1" & vFormat & ".asp"     _
         & "?vStrDate="        & Server.UrlEncode(vStrDate) _
         & "&vCurList="        & vCurList       _
         & "&vFind="           & vFind          _
         & "&vFindId="         & vFindId        _
         & "&vFindFirstName="  & vFindFirstName _
         & "&vFindLastName="   & vFindLastName  _
         & "&vFindCriteria="   & vFindCriteria  _
         & "&vFindEmail="      & vFindEmail
    Response.Redirect vUrl
'   Response.Write vUrl

  End If
  
%>
<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <title></title>
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <table border="0" width="100%" cellspacing="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111">
    <tr>
      <td valign="top">
      <h1 align="center">Passport to Safety Report Card</h1>
      <h2>This displays learners who have passed the Passport to Safety Course.&nbsp; You can either display the report online or create an <i>MS Excel</i> spreadsheet which can be saved on your <i>Desktop</i> for analysis. <font color="#FF0000">Please be patient, the Excel format can take several minutes.</font></h2>
      </td>
    </tr>
    <tr>
      <td valign="top">
      <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript" Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vFindCriteria.selectedIndex < 0)
  {
    alert("Please select one of the \"College | Faculty\" options.");
    theForm.vFindCriteria.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="ReportCard_1.asp" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript" name="FrontPage_Form1">
        <table border="0" width="100%" cellpadding="2" bordercolor="#DDEEF9" style="border-collapse: collapse">
          <tr>
            <th align="right" valign="top">Select College | Faculty:</th>
            <td>
              <% i =  fCriteriaList (svCustAcctId, "KIDS:" & svMembCriteria) '...huge forces managers to be treated like facs %>&nbsp;&nbsp; <!--webbot bot="Validation" s-display-name="College | Faculty" b-value-required="TRUE" -->
              <select size="<%=vCriteriaListCnt%>" name="vFindCriteria" multiple><%=i%></select>
            </td>
          </tr>
          <tr>
            <th align="right" nowrap valign="top">
            <%    
              vOption = ""
              For i = 1 To 9
                Select Case i
                  Case 1 : j =  1  : vDesc = "<!--{{-->1 day<!--}}-->"
                  Case 2 : j =  7  : vDesc = "7 " & "<!--{{-->days<!--}}-->"
                  Case 3 : j = 14  : vDesc = j & " " & "<!--{{-->days<!--}}-->"
                  Case 4 : j = 30  : vDesc = j & " " & "<!--{{-->days<!--}}-->"
                  Case 5 : j = 60  : vDesc = j & " " & "<!--{{-->days<!--}}-->"
                  Case 6 : j = 90  : vDesc = j & " " & "<!--{{-->days<!--}}-->"
                  Case 7 : j = 180 : vDesc = j & " " & "<!--{{-->days<!--}}-->"
                  Case 8 : j = 365 : vDesc = j & " " & "<!--{{-->days<!--}}-->"
                  Case 9 : j = 999 : vDesc = "<!--{{-->all available days<!--}}-->"
                End Select
                k = fFormatSqlDate(DateAdd ("d", -j, Now))
                vSelected = fIf(vStrDate = k, " selected", "")
                vOption = vOption & "<option value='" & k & "'" & vSelected & ">" & vDesc & "</option>" & vbCrLf 
              Next
            %> 
            Completed within :
            </th>
            <td><select size="1" name="vStrDate"><%=vOption%></select></td>
          </tr>
          <tr>
            <th align="right" valign="top">Selecting learners :</th>
            <td><input type="radio" name="vFind" value="S" <%=fcheck("s", vfind)%>>starting with <br><input type="radio" name="vFind" value="C" <%=fcheck("c", vfind)%>>containing</td>
          </tr>
          <tr>
            <td align="right" valign="top">&nbsp;Learner ID :</td>
            <td><input type="text" name="vFindId" size="29" value="<%=vFindId%>"></td>
          </tr>
          <tr>
            <td align="right" valign="top">First Name :</td>
            <td><input type="text" name="vFindFirstName" size="29" value="<%=vFindFirstName%>">&nbsp; </td>
          </tr>
          <tr>
            <td align="right" valign="top">Last Name :</td>
            <td><input type="text" name="vFindLastName" size="29" value="<%=vFindLastName%>"></td>
          </tr>
          <tr>
            <td align="right" valign="top">Email Address :</td>
            <td><input type="text" name="vFindEmail" size="29" value="<%=vFindEmail%>"></td>
          </tr>
          <tr>
            <th nowrap width="100%" height="50" colspan="2">
            <h2><br>Then click either....</h2>
            <input type="submit" value="Display Online" name="bPrint" class="button"> or 
            <input type="submit" value="MS Excel File" name="bExcel" class="button">
            <p>&nbsp;</p>
            </th>
          </tr>
        </table>
        <input type="hidden" name="vForm" value="y">
      </form>
      </td>
    </tr>
  </table>
  <!--#include virtual = "V5\Inc\Shell_Lo.asp"-->

</body>

</html>