<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Routines.asp"-->

<%
  Dim vDat, vCou, vGrs, vPrs, aPrs, vPrv, vHst, vGst, vPst, vBg
  vCou = fDefault(Request("vCou"), "CA")
  vDat = fDefault(Request("vDat"), "Jul 01, 2010")
  vGrs = fDefault(Request("vGrs"), 100)
  vPrs = "AB BC MB NB NF NT NS NU ON PE QC SK YT"
%>

<html>

<head>
  <meta http-equiv="Content-Language" content="en-us">
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <title>VUBIZ Tax Calculator</title>
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0">

  <form method="POST" action="Taxes_Test.asp">
  <div align="center">
  <table border="0" bordercolor="#DCEDF8" cellpadding="10" style="border-collapse: collapse;" class="table" width="800">
      <tr>
        <td width="100%" align="center" class="cell_U1" style="height: 130px">
          <h3>Vubiz Tax Calculator</h3>
          <p>Show tax on $: 
            <input type="text" name="vGrs" size="6" value="100" maxlength="6" class="c2">&nbsp; on 
            <select size="1" name="vDat" class="c2">
              <option value="Dec 31, 2007" <%=fSelect("Dec 31, 2007", vDat)%>>Dec 31, 2007</option>
              <option value="Jan 01, 2008" <%=fSelect("Jan 01, 2008", vDat)%>>Jan 01, 2008</option>
              <option value="Jul 01, 2010" <%=fSelect("Jul 01, 2010", vDat)%>>Jul 01, 2010</option>
            </select>&nbsp; in             
            <input type="radio" value="CA" <%=fCheck("CA", vCou)%> name="vCou">CA
            <input type="radio" value="US" <%=fCheck("US", vCou)%> name="vCou">Other
            <input type="submit" value="Submit" name="B3" class="button">
          </p>
        </td>
      </tr>
      <tr>
        <td>
        <div align="center">
          <table border="0" bordercolor="#DCEDF8" cellpadding="5" style="border-collapse: collapse; width:90%" class="table">
            <tr>
              <th>Prov</th>
              <th align="right">$Gross</th>
              <th align="right">$HST</th>
              <th align="right">$GST</th>
              <th align="right">$PST</th>
              <th align="right">$Total</th>
            </tr>
             <%

              aPrs = Split(vPrs)
              For i = 0 To Ubound(aPrs)

                vHst = vGrs * fHST (vDat, vCou, aPrs(i))
                vGst = vGrs * fGST (vDat, vCou, aPrs(i))
                vPst = vGrs * fPST (vDat, vCou, aPrs(i))
                vBg = fIf (i Mod 2 = 0, "", " bgcolor='#DCEDF8'")
                If aPrs(i) = "ON" Then vBg = " bgcolor='#FFFF00'"
%>
            <tr>
              <th <%=vBg%>><%=aPrs(i)%></th>
              <td align="right" <%=vBg%>><%=FormatNumber(vGrs, 2)%></td>
              <td align="right" <%=vBg%>><%=FormatNumber(vHst, 2)%></td>
              <td align="right" <%=vBg%>><%=FormatNumber(vGst, 2)%></td>
              <td align="right" <%=vBg%>><%=FormatNumber(vPst, 2)%></td>
              <td align="right" <%=vBg%>><%=FormatNumber(vGrs+vHst+vGst+vPst, 2)%></td>
            </tr>
<%
            Next
%>
          </table>
        </div>
        </td>
      </tr>
    <tr>
      <td align="center" class="cell_L1" colspan="4">&nbsp;</td>
    </tr>
  </table>
  </div>
  </form>

</body>

</html>