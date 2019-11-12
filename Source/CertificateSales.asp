<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<%
  Dim vRecNo, vTypes, vCerts, aCerts, vYears, vScript, aCertsAll, aYearsAll
  Dim vProg, vCert, vCertPrev, vChannel
  Dim vDate, vM, vD, vY, vCA0, vUS0, vCA1, vUS1, vCnt1, vCA2, vUS2, vCnt2, vCrit, vCritPrev, vCntEN1, vCntFR1, vCntEN2, vCntFR2

  vTypes = fDefault(Request("vTypes"), "rs")
  vCerts = fDefault(Request("vCerts"), "*")
' vYears = fDefault(Request("vYears"), Right(Year(Now), 2))
  vYears = fDefault(Request("vYears"), Year(Now))

  If vCerts <> "*" Then vCerts = "'" & Replace(vCerts, ", ", "','") & "'"
  If vYears <> "*" Then vYears = "'" & Replace(vYears, ", ", "','") & "'"

  vRecNo = 0  '...this counts number of records for the report header, etc
  
  '...Certs dropdown
  Function fCerts (vCerts)
    Dim vSelected, vAllCerts
    vAllCerts = ""
    fCerts = ""
    sOpenDb
    vSql = "SELECT * FROM Cert "
    Set oRs = oDb.Execute(vSql)    
    Do While Not oRs.Eof
      vSelected = fIf(Instr(vCerts, oRs("Cert_Id")) > 0 , " Selected ", "")
      fCerts = fCerts & "<option value='" & oRs("Cert_Id") & "'" & vSelected & ">" & oRs("Cert_Id") & "</option>" & vbCRLF
      oRs.MoveNext
    Loop
    Set oRs = Nothing
    sCloseDb           
  End Function    

   '...Years dropdown
  Function fYears (vYears)
    Dim vSelected, vYear
    fYears = ""    
'   For i = 2006 To Year(Now)
    For vYear = 2006 To Year(Now)
'     vYear = Right(i, 2)
      vSelected = fIf(Instr(vYears, vYear) > 0 , " Selected ", "")
'     fYears = fYears & "<option value='" & vYear & "'" & vSelected & ">" & i & "</option>" & vbCRLF
      fYears = fYears & "<option value='" & vYear & "'" & vSelected & ">" & vYear & "</option>" & vbCRLF
    Next
  End Function


  '...Get Certs and Years for Chart 
  Sub sGetChartNeeds
    Dim i  
 
    If vCerts = "*" Then
      sOpenDb
      vSql = "SELECT DISTINCT Cert_Id FROM Cert ORDER BY Cert_Id"
      Set oRs = oDb.Execute(vSql)    
      i = ""
      Do While Not oRs.Eof
        i = i & oRs("Cert_Id") & "|" 
        oRs.MoveNext
      Loop
      Set oRs = Nothing
      sCloseDb      
      aCertsAll = Split(Left(i, Len(i)-1), "|")
    Else  
      i = Mid(vCerts, 2, Len(vCerts) - 2)
      i = Replace(i, "','", "|")
      aCertsAll = Split(i, "|")
    End If

    If vYears = "*" Then
      i = ""
      For j = 2006 To Year(Now)
        i = i & j & " "
      Next
      aYearsAll = Split(Trim(i))
    Else  
      i = Mid(vYears, 2, Len(vYears) - 2)
      i = Replace(i, "','", "|")
      aYearsAll = Split(i, "|")
      For j = 0 To Ubound(aYearsAll) '...drop down only uses last two chars of year
        aYearsAll(j) = "20" & aYearsAll(j)
      Next
    End If
  End Sub

%>

<html>

<head>
  <title>CertificateSales</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <style>
    #mainReport { width: 700px; }

    td { text-align: center; }
  </style>

</head>
<body>

  <% Server.Execute vShellHi %>

  <h1>Certificate Sales Report</h1>
  <h2>This reports Certificate Programs sold by Partners (4 characters) and Channels (8 characters) via Ecommerce.</h2>

  <form method="POST" action="CertificateSales.asp">
    <div style="text-align: center">
      <table>
        <tr>
          <td style="text-align: right">
            <select size="14" name="vCerts" multiple style="width: 120px; margin: 10px;">
              <option <%=fselects(vcerts, "*")%> value="*">All Certificates</option>
              <%=fCerts(vCerts)%>
            </select>
          </td>
          <td>
            <select size="14" name="vYears" multiple style="width: 120px; margin: 10px;">
              <option <%=fselects(vYears, "*")%> value="<%="*"%>">All Years</option>
              <%=fYears(vYears)%>
            </select>
          </td>
        </tr>
        <tr>
          <td colspan="2" style="text-align: center">Use Ctrl+Enter for multiple selections (except for &quot;All&quot;).<br />
            <br />
          </td>
        </tr>
        <tr>
          <td style="text-align: center" colspan="2">
            <input type="radio" value="rs" name="vTypes" <%=fCheck("rs", vTypes)%>>Summary Report<%=f10%>
            <input type="radio" value="rd" name="vTypes" <%=fCheck("rd", vTypes)%>>Detailed Report 
          </td>
        </tr>
        <tr>
          <td colspan="2" style="text-align: center">
            <br />
            <input type="submit" value="Go" name="bGo" class="button" onclick="$(this).hide()">
          </td>
        </tr>
      </table>
    </div>
  </form>

  <%  
    vCritPrev = ""
    vCA1      = 0
    vUS1      = 0
    vCnt1     = 0
    vCnt2     = 0
    vCA2      = 0
    vUS2      = 0
    vCntEN1   = 0
    vCntFR1   = 0
    vCntEN2   = 0
    vCntFR2   = 0
    
    sOpenDb

    '...delete any existing records
    vSql = "" _
         & "DELETE Cert_Temp WHERE UserNo = " & svMembNo
'        sDebug
         sOpenDb
         oDb.Execute(vSql)

    '...Get Special Certs and join with the Channel Certs - diff is in the length (Type)
    '   as they are being selected, put them into a charting table (CerC) [ first drop it if it exists ]

    vSql = "" _
         & "INSERT INTO Cert_Temp " _

         & "SELECT DISTINCT "_ 
         & "  Cert.Cert_Order, "_ 
         & "  Cert.Cert_Id, "_ 
         & "  Memb.Memb_AcctId, "_ 
         & "  COALESCE(Crit.Crit_Id, 'NONE') AS Source, "_ 
         & "  Ecom.Ecom_CardName, "_ 
         & "  Ecom.Ecom_Prices, "_ 
         & "  Ecom.Ecom_Currency, "_ 
         & "  Ecom.Ecom_Media, "_ 
         & "  Ecom.Ecom_Programs, "_
         & "  4 AS [Type], "_
         &    svMembNo & " AS [UserNo], "_ 
         & "  Ecom.Ecom_Issued "_ 
         & "FROM "_
         & "  Memb RIGHT OUTER JOIN "_
         & "  Cert ON CHARINDEX(Memb.Memb_AcctId, Cert_Accts) > 0 LEFT OUTER JOIN"_
         & "  Crit ON Memb.Memb_Criteria = Crit.Crit_No INNER JOIN "_
         & "  Ecom ON SUBSTRING(Memb.Memb_Memo, 7, 99) = Ecom.Ecom_OrderNo "_
         & "WHERE "_
         & "  (Ecom.Ecom_Media = 'Spec_01') AND "_
         & "  (Ecom.Ecom_Prices <> 0) AND "_
         & "  (ISNUMERIC(COALESCE(Memb.Memb_Criteria, 0)) = 1) "_
         &    fIf(vCerts <> "*", "AND (Cert.Cert_Id          IN (" & vCerts & ")) ", "") _
         &    fIf(vYears <> "*", "AND (YEAR(Ecom_Issued) IN (" & vYears & ")) ", "") _

         & "UNION "_

         & "SELECT DISTINCT "_
         & "  Cert.Cert_Order, "_ 
         & "  Cert.Cert_Id, "_ 
         & "  Ecom.Ecom_AcctId, "_ 
         & "  Ecom.Ecom_CustId AS [Source], "_ 
         & "  Ecom.Ecom_CardName, "_ 
         & "  Ecom.Ecom_Prices, "_ 
         & "  Ecom.Ecom_Currency, "_ 
         & "  Ecom.Ecom_Media, "_ 
         & "  Ecom.Ecom_Programs, "_
         & "  8 AS [Type], "_
         &    svMembNo & " AS [UserNo], "_ 
         & "  Ecom.Ecom_Issued "_ 
         & "FROM "_
         & "  Ecom INNER JOIN"_
         & "  Cert ON CHARINDEX(Left(Ecom_Programs, 5), Cert_Progs) > 0 "_
         & "WHERE "_
         & "  (Ecom.Ecom_Prices <> 0) "_
         &    fIf(vCerts <> "*", "AND (Cert.Cert_Id IN (" & vCerts & ")) ", "") _
         &    fIf(vYears <> "*", "AND (YEAR(Ecom_Issued) IN (" & vYears & ")) ", "") _

         & "ORDER BY "_
         & "  Cert_Order, "_
         & "  Type, "_
         & "  Source, "_
         & "  Ecom.Ecom_Issued "
'   sDebug
    sOpenDb
    oDb.Execute(vSql)
    sCloseDb


    sOpenDb
    vSql = "SELECT * From Cert_Temp WHERE UserNo = " & svMembNo & " ORDER BY Cert_Order, Type, Source"

    Set oRs = oDb.Execute(vSql)
    Do While Not oRs.Eof    
      vRecNo = vRecNo + 1
      sDisplay
      oRs.MoveNext
    Loop
    Set oRs = Nothing 
    sCloseDb
    If vRecNo = 0 Then 
      sNoData 
    Else
      sFinish
    End If


  Sub sDisplay    
    '...get the partner from the criteria
    If oRs("Ecom_Media") = "Spec_01" Then
      vCrit = oRs("Source") '...might be NONE when no partners, like ICSA
    '...get the channel from the cust id/channel
    Else
      vCrit = oRs("Source")
    End If              
    
    If vRecNo = 1 Then 
  %>

  <table id="mainReport" style="margin: auto;">
    <%
    End If

    If vCritPrev <> "" And vCrit <> vCritPrev Then
    %>
    <tr>
      <td style="text-align: left;" class="ro"><%=vCertPrev%></td>
      <td style="text-align: left; font-weight:bold" colspan="3"><%=vCritPrev%></td>
      <td><%=vCntEN1%></td>
      <td><%=vCntFR1%></td>
      <td><%=vCnt1%></td>
      <td><%=FormatCurrency(vCA1)%></td>
      <td><%=FormatCurrency(vUS1)%></td>
    </tr>
    <%  If vTypes = "rd" Then %>
    <tr>
      <td colspan="9">&nbsp;</td>
    </tr>
    <% 
      End If 

      vCA1    = 0
      vUS1    = 0
      vCnt1   = 0
      
      vCntEN1 = 0
      vCntFR1 = 0
    End If

    vCA0 = 0
    vUS0 = 0

    vCnt1 = vCnt1 + 1
    vCnt2 = vCnt2 + 1

    If oRs("Ecom_Currency") = "CA" Then
      vCA0 = oRs("Ecom_Prices")
      vCA1 = vCA1 + oRs("Ecom_Prices")
      vCA2 = vCA2 + oRs("Ecom_Prices")
    Else
      vUS0 = oRs("Ecom_Prices")
      vUS1 = vUS1 + oRs("Ecom_Prices")
      vUS2 = vUS2 + oRs("Ecom_Prices")
    End If
    
    If Right(oRs("Ecom_Programs"), 2) = "EN" Then
      vCntEN1 = vCntEN1 + 1
      vCntEN2 = vCntEN2 + 1
    Else
      vCntFR1 = vCntFR1 + 1
      vCntFR2 = vCntFR2 + 1
    End If

    '... this is the ecommerce format for the date sold: "0610-0316-3526"
'   vDate = oRs("Ecom_OrderNo")
    vDate = oRs("Ecom_Issued")
'    vY = 2000 + Mid(vDate, 1, 2)
'    vM = fFormatMonth (Mid(vDate, 3, 2))
'    vD = Mid(vDate, 6, 2)
'    vDate = vM & " " & vD & ", " & vY
    
    vCert = oRs("Cert_Id")
    vCertPrev = vCert
    vCritPrev = vCrit

    If vRecNo = 1 Then sHeader

    If vTypes = "rd" Then
    %>
    <tr>
      <td style="text-align: left;"><%=vCert%></td>
      <td style="text-align: left;"><%=vCrit%></td>
      <td><%=fFormatDate(vDate)%></td>
      <td style="text-align: left"><%=oRs("Ecom_CardName")%></td>
      <td><%=fIf(Right(oRs("Ecom_Programs"), 2) = "EN", "<img border='0' src='../Images/Icons/CheckMark.jpg' width='12' height='15'", "")%></td>
      <td><%=fIf(Right(oRs("Ecom_Programs"), 2) = "FR", "<img border='0' src='../Images/Icons/CheckMark.jpg' width='12' height='15'", "")%></td>
      <td></td>
      <td><%=FormatCurrency(vCA0)%></td>
      <td><%=FormatCurrency(vUS0)%></td>
    </tr>

    <%
      End If

    End Sub 
    %>


    <% Sub sHeader%>
    <tr>
      <th class="rowshade">Certificate</th>
      <th style="text-align: left;" class="rowshade">Channel</th>
      <th class="rowshade"><%=fIf(vTypes="rd","Date", "")%></th>
      <th class="rowshade" style="text-align: left"><%=fIf(vTypes="rd","Cardholder", "")%></th>
      <th class="rowshade">#EN</th>
      <th class="rowshade">#FR</th>
      <th class="rowshade">#Total</th>
      <th class="rowshade">$ CA</th>
      <th class="rowshade">$ US</th>
    </tr>
    <% End Sub %>


    <% Sub sFinish %>
    <tr>
      <td style="text-align: left;"><%=vCertPrev%></td>
      <td style="text-align: left; font-weight: bold" colspan="3"><%=vCritPrev%></td>
      <td><%=vCntEN1%></td>
      <td><%=vCntFR1%></td>
      <td><%=vCnt1%></td>
      <td><%=FormatCurrency(vCA1)%></td>
      <td><%=FormatCurrency(vUS1)%></td>
    </tr>
    <tr>
      <td colspan="9">&nbsp;</td>
    </tr>
    <tr>
      <td></td>
      <td style="text-align: left; font-weight: bold" colspan="3">Total</td>
      <td><%=vCntEN2%></td>
      <td><%=vCntFR2%></td>
      <td><%=vCnt2%></td>
      <td><%=FormatCurrency(vCA2)%></td>
      <td><%=FormatCurrency(vUS2)%></td>
    </tr>
  </table>
  <% End Sub %>


  <% Sub sNoData %>
  <h6>
    <br>
    <br>
    <br>
    There is no data available for that Certificate/Year selection.<br>
    <br>
    <br>
  </h6>
  <% End Sub %>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
