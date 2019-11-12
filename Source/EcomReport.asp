<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Prod.asp"-->
<!--#include virtual = "V5/Inc/Db_Cont.asp"-->

<% 
  '...ensure users and/or facilitators don't try to run this report by bypassing the menu page
  If svMembLevel < 4 Then Response.Redirect "Menu.asp"

  Dim vCustIdPrev, vPrograms, vNameInfo, vAddressInfo, vTotProgs, vTotRefs, vStrDate, vEndDate, vStrDateErr, vEndDateErr, vButton, vColor, vFreebie, vWhere

  '...split values: owner %s come from prod table and cust % comes from the customer table
  '   for the admin report, aOwnr build summaries
  Dim vEcom_SplitVubz, vEcom_SplitCust, vEcom_SplitOwnr, aOwnr_CA(), aOwnr_US(), vOwnrCnt
    
  Dim vSplitVubz, vTotSplitVubz_US, vTotSplitVubz_CA, vGrandTotSplitVubz_US, vGrandTotSplitVubz_CA
  Dim vSplitOwnr, vTotSplitOwnr_US, vTotSplitOwnr_CA, vGrandTotSplitOwnr_US, vGrandTotSplitOwnr_CA
  Dim vSplitCust, vTotSplitCust_US, vTotSplitCust_CA, vGrandTotSplitCust_US, vGrandTotSplitCust_CA  
    
  Dim vAmount, vTotAmount_US, vTotAmount_CA, vGrandTotAmount_US, vGrandTotAmount_CA
  Dim vPrice, vTotPrice_US, vTotPrice_Ca, vGrandTotPrice_US, vGrandTotPrice_CA
  Dim vProgram, vOrderId

  Dim vAddress, vMthStr, vMthEnd, vDate, vDateUrl, vOption1, vOption2, vDateMonth, vSelected, vExpires, vIssued, vCurrDate, vOwnerId, vSource
  Dim vTotPST, vTotGST, vTotHST, vTotTax

  vPrograms = Trim(Request("vPrograms"))
  vAddress = Request("vAddress")
  vOwnerId = Ucase(Request("vOwnerId"))
  vFreebie = Request("vFreebie")

  '...get source of posting
  vSource = ""
  If fDefault(Request("vSource_E"), "Y") = "Y" Then vSource = "E"
  If fDefault(Request("vSource_V"), "N") = "Y" Then vSource = vSource & "V"
  If fDefault(Request("vSource_C"), "N") = "Y" Then vSource = vSource & "C"

  '...if Excel then go to Excel version
  vButton = "Online Report"
  If Request.Form("bExcel").Count = 1 Then 
    Response.Redirect "EcomReportX.asp?vStrDate=" & Server.UrlEncode(Request("vStrDate")) & "&vEndDate=" & Server.UrlEncode(Request("vEndDate")) & "&vPrograms=" & Request("vPrograms") & "&vAddress=" & Request("vAddress") & "&vOwnerId=" & Request("vOwnerId") & "&vFreebie=" & Request("vFreebie") & "&vSource=" & vSource
  End If

  '...defaults to current month
  If Request("vStrDate").Count = 0 And Request("vEndDate").Count = 0 Then
    vStrDate  = Request("vStrDate")      : If Len(vStrDate) = 0 Then vStrDate = fFormatDate(MonthName(Month(Now)) & " 1, " & Year(Now))
    vEndDate  = Request("vEndDate")      : If Len(vEndDate) = 0 Then vEndDate = fFormatDate(DateAdd("d", -1, MonthName(Month(DateAdd("m", +1, Now))) & " 1, " & Year(DateAdd("m", +1, Now))))
  Else
    vStrDate  = fFormatDate(Request("vStrDate")) 
    If Request("vStrDate") = "" Then 
      vStrDate = ""
    ElseIf vStrDate = " " Then
      vStrDate  = Request("vStrDate") 
      vStrDateErr = "Error"
    End If
    vEndDate  = fFormatDate(Request("vEndDate"))
    If Request("vEndDate") = "" Then 
      vEndDate = ""
    ElseIf vEndDate = " " Then
      vEndDate  = Request("vEndDate") 
      vEndDateErr = "Error"
    End If
    If (Len(vStrDate) > 0 And vStrDateErr = "") And (Len(vEndDate) > 0 And vEndDateErr = "") Then
      If DateDiff("d", vStrDate, vEndDate) < 0 Then
        vEndDateErr = "Error"
      End If
    End If
  End If


  '...This removes the html blank enabling better column appearance
  Function fRemoveBlank(x)
    fRemoveBlank = x
    If Len(x) > 6 Then
      If Right(x, 6) = "&nbsp;" Then
'       fRemoveBlank = Left(x, len(x)-6)
      End If
    End If
  End Function
  

  Function fMedia
    Select Case vEcom_Media
      Case "CDs"       : fMedia = "CD "
      Case "Prods"     : fMedia = "PR "
      Case "Group"     : fMedia = "G1 "
      Case "Group2"    : fMedia = "G2 "
      Case "AddOn2"    : fMedia = "G2 "
      Case "Spec_01"   : fMedia = "S1 "
      Case Else        : fMedia = "IO "
    End Select
  End Function
 
%>

<html>

<head>
  <title>EcomReport.asp</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>


  <style>
    .table .d1 td:nth-child(01) { text-align: left; }
    .table .d1 td:nth-child(02) { text-align: left; }
    .table .d1 td:nth-child(03) { text-align: center; white-space:nowrap; }
    .table .d1 td:nth-child(04) { text-align: center; white-space:nowrap; }
    .table .d1 td:nth-child(05) { text-align: right; }
    .table .d1 td:nth-child(06) { text-align: right; }
    .table .d1 td:nth-child(07) { text-align: right; }
    .table .d1 td:nth-child(08) { text-align: right; }
    .table .d1 td:nth-child(09) { text-align: right; }

    .table .t1 td:nth-child(01) { width:22%; text-align: left;  }
    .table .t1 td:nth-child(02) { width:22%; text-align: left; }
    .table .t1 td:nth-child(03) { width:08%; text-align: center; }
    .table .t1 td:nth-child(04) { width:08%; text-align: center; }
    .table .t1 td:nth-child(05) { width:08%; text-align: center; }
    .table .t1 td:nth-child(06) { width:08%; text-align: center; }
    .table .t1 td:nth-child(07) { width:08%; text-align: center; }
    .table .t1 td:nth-child(08) { width:08%; text-align: center; }
    .table .t1 td:nth-child(09) { width:08%; text-align: center; }

    .table .h1 td:nth-child(01) { font-weight:bold; text-align: left; }
    .table .h1 td:nth-child(02) { font-weight:bold; text-align: right; }
    .table .h1 td:nth-child(03) { font-weight:bold; text-align: right; }
    .table .h1 td:nth-child(04) { font-weight:bold; text-align: right; }
    .table .h1 td:nth-child(05) { font-weight:bold; text-align: right; }
    .table .h1 td:nth-child(06) { font-weight:bold; text-align: right; }
  </style>



</head>

<body>

  <% Server.Execute vShellHi %>


  <h1>Basic Ecommerce Sales Report</h1>
  <h3>
    <%=fIf(IsDate(vStrDate), vStrDate, "First Sale")%> - <%=fIf(IsDate(vEndDate), vEndDate, "Last Sale")%><br /><br />
    <%=fIf(Len(vOwnerID) > 0, "(Includes sales of Owner Id: " & vOwnerId & ")", "")  %>&nbsp;&nbsp;<a href="#" onclick="toggle('divDesc')">[Click for details]</a><br />
  </h3>

  <div class="div" id="divDesc" style="padding: 20px; width: 80%; margin: auto;">
    If appropriate, enter Owner Id code to include all ecommerce transactions that have been made by other accounts/channels for proprietary content.&nbsp; If this field is blank, then only content sold via this account/channel will be displayed.&nbsp; &quot;Include addresses&quot; displays the customer&#39;s full address information (for your channel only). Selecting month further narrows down the view.&nbsp;
    <br /><br />The Program Code (in green if manually created/adjusted) is followed by an &quot;E&quot; (normal ecommerce), &quot;M&quot; (manual payments to customer) or &quot;V&quot; (manual payments to Vubiz) followed by an &quot;IO&quot; (individual online), &quot;G1&quot; or &quot;G2&quot; (group online).
    <br /><br />Column &quot;%Own&quot; is the percentage revenue to the Content's Owner.. Column &quot;%Chn&quot; is the percentage revenue to the Channel Reseller (calculated after the Owner's percentage has been deducted).&nbsp; Column &quot;Price&quot; is the dollar amount of the sale.&nbsp; Column &quot;$Total+&quot; contains taxes and shipping.&nbsp; Note that detailed &quot;split&quot; values have been rounded but Totals are accurate.
  </div>


  <table class="table">

    <!--  This starts the section that is created if the form has not been filled in ok -->
    <% 
        '...If first pass then display the drop down form
        If Request.Form("vHidden").Count = 0 Or vStrDateErr <> "" Or vEndDateErr <> "" Then
    %>

    <tr>
      <td colspan="9">
        &nbsp;
        
          <form method="POST" action="EcomReport.asp">
            <input type="Hidden" name="vHidden" value="Hidden">

            <table>

              <tr>
                <th>Include sales by Owner Id :</th>
                <td>
                  <input type="text" name="vOwnerId" size="8" value="<%=vOwnerId%>"><br>If you are an author/owner, enter your Owner Id and this report will include other accounts that have sold your content.&nbsp; Leave empty to see sales for this account only.
                </td>
              </tr>
              <tr>
                <th>For Program Id(s) :</th>
                <td>
                  <input type="text" name="vPrograms" size="59" value="<%=vPrograms%>">
                  <br>ie P1234EN P2345EN. Leave emtpy for all programs. Separate multiple Program IDs with a space.
                </td>
              </tr>
              <tr>
                <th>Include addresses ?</th>
                <td>
                  <input type="checkbox" name="vAddress" value="Y" <%=fcheck("y", vaddress)%>>Tick for Yes (only available for this account)
                </td>
              </tr>
              <tr>
                <th>Include free programs ?</th>
                <td>
                  <input type="checkbox" name="vFreebie" value="Y" <%=fcheck("y", vfreebie)%>>Tick for Yes (Otherwise, programs tagged as &quot;No Charge&quot; will not be displayed).
                </td>
              </tr>
              <tr>
                <th>Select Start Date :</th>
                <td>
                  <input type="text" name="vStrDate" size="15" value="<%=vStrDate%>">
                  <span style="background-color: #FFFF00"><%=vStrDateErr%></span><br>ie Jan 1, 2006 (MMM DD, YYYY). Leave empty to start at first sale.
                </td>
              </tr>
              <tr>
                <th>Select End Date :</th>
                <td>
                  <input type="text" name="vEndDate" size="15" value="<%=vEndDate%>">
                  <span style="background-color: #FFFF00"><%=vEndDateErr%></span><br>ie Mar 31, 2006 (MMM DD, YYYY). Leave empty to finish with last sale.
                </td>
              </tr>
              <tr>
                <th>Include Transactions : </th>
                <td>
                  <input type="radio" value="Y" name="vSource_E" <%=fcheck("y", fdefault(request("vsource_e"), "y"))%>>Yes&nbsp;&nbsp; 
                <input type="radio" value="N" name="vSource_E" <%=fcheck("n", fdefault(request("vsource_e"), "y"))%>>No&nbsp;&nbsp; - Normal E-ecommerce<br>
                  <input type="radio" value="Y" name="vSource_V" <%=fcheck("y", fdefault(request("vsource_v"), "y"))%>>Yes&nbsp;&nbsp; 
                <input type="radio" value="N" name="vSource_V" <%=fcheck("n", fdefault(request("vsource_v"), "y"))%>>No&nbsp;&nbsp; - Manual Payment to Vubiz<br>
                  <input type="radio" value="Y" name="vSource_C" <%=fcheck("y", fdefault(request("vsource_c"), "y"))%>>Yes&nbsp;&nbsp; 
                <input type="radio" value="N" name="vSource_C" <%=fcheck("n", fdefault(request("vsource_c"), "y"))%>>No&nbsp;&nbsp; - Manual Payment to Channel
                </td>
              </tr>
              <tr>
                <td colspan="2" style="text-align: center; padding: 20px;">
                  <input type="submit" value="Online" name="bPrint" id="bPrint" class="button">&ensp;&ensp;&ensp;
                <input type="submit" value="Excel" name="bExcel" class="button">
                </td>
              </tr>
            </table>
          </form>

      </td>
    </tr>
    <!--  This ends the section that is created if the form has not been filled in ok -->
    <!--  This starts the section that is created if the form is filled in ok -->
    <% Else %>
    <tr class="t1">
      <th class="rowshade" style="text-align: left;">Customer</th>
      <th class="rowshade" style="text-align: left;">Program</th>
      <th class="rowshade">Issued</th>
      <th class="rowshade">Expires</th>
      <th class="rowshade">$Vubiz&nbsp;&nbsp; </th>
      <th class="rowshade">$Cust&nbsp; </th>
      <th class="rowshade">$Owner&nbsp;&nbsp; </th>
      <th class="rowshade">$Total&nbsp;&nbsp; </th>
      <th class="rowshade">$Total+&nbsp;&nbsp; </th>
    </tr>
    <%

      vCustIdPrev            = ""
      vTotProgs              = 0
      vTotRefs               = 0
      vOwnrCnt               = 0
  
      Redim Preserve aOwnr_CA(vOwnrCnt)
      Redim Preserve aOwnr_US(vOwnrCnt)
  
      vTotSplitVubz_US       = 0
      vTotSplitVubz_CA       = 0
      vGrandTotSplitVubz_US  = 0
      vGrandTotSplitVubz_CA  = 0
      vTotSplitOwnr_US       = 0
      vTotSplitOwnr_CA       = 0
      vGrandTotSplitOwnr_US  = 0 
      vGrandTotSplitOwnr_CA  = 0
      vTotSplitCust_US       = 0
      vTotSplitCust_CA       = 0
      vGrandTotSplitCust_US  = 0 
      vGrandTotSplitCust_CA  = 0
      vTotPrice_US           = 0
      vTotPrice_CA           = 0
      vGrandTotPrice_US      = 0
      vGrandTotPrice_CA      = 0
      vTotAmount_US          = 0
      vTotAmount_CA          = 0
      vGrandTotAmount_US     = 0
      vGrandTotAmount_CA     = 0

      vTotPST                = 0
      vTotGST                = 0
      vTotHST                = 0
      vTotTAX                = 0
  
      '...access restriction rules (create the "WHERE" part of the sql statement)
      '   administrators can only see all accounts (ie need to be a manager to just see ur own stuff)
      '   owners can see all accounts with their owner id
      '   rest just see their account
      '   anyone can select a month or all months

      vWhere = " WHERE "

      vSql = " SELECT * FROM Ecom Ec WITH (nolock) " _
           & "   LEFT OUTER JOIN Cust Cu WITH (nolock) ON Ec.Ecom_CustId = Cu.Cust_Id " _
           & "   LEFT OUTER JOIN Memb Me WITH (nolock) ON Ec.Ecom_MembNo = Me.Memb_No " 

      If Len(vOwnerId) = 4 Then
        vSql = vSql & " LEFT OUTER JOIN V5_Base.dbo.Prog Pr WITH (nolock) ON (Ec.Ecom_Programs = Pr.Prog_Id) AND (Pr.Prog_Owner = '" & vOwnerId & "') "
      End If

      If Len(vOwnerId) = 4 Then
        If svMembLevel = 4 And Not svMembManager Then
          vWhere = vWhere & " ((Cu.Cust_AcctId = '" & svCustAcctId & "') OR (Pr.Prog_Owner = '" & vOwnerId & "')) AND "       
        ElseIf svMembLevel = 4 And svMembManager Then
          vWhere = vWhere & " ((Cu.Cust_Id LIKE '" & Left(svCustId, 4) & "%')  OR (Pr.Prog_Owner = '" & vOwnerId & "')) AND "
        End If
      Else
        If svMembLevel = 4 And Not svMembManager Then
          vWhere = vWhere & "(Cu.Cust_AcctId = '" & svCustAcctId & "') AND "       
        ElseIf svMembLevel = 4 And svMembManager Then
          vWhere = vWhere & " (Cu.Cust_Id LIKE '" & Left(svCustId, 4) & "%') AND "
        End If
      End If

      If Len(vStrDate) > 6  Then vWhere = vWhere & " (Ecom_Issued >= '" & vStrDate & "') AND "
      If Len(vEndDate) > 6  Then vWhere = vWhere & " (Ecom_Issued < DATEADD(d, 1, '" & vEndDate & "')) AND " 

      If vFreebie <> "Y"    Then vWhere = vWhere & " (Ecom_Amount <> 0) AND "
			If Len(vPrograms) > 0 Then vWhere = vWhere & " (CHARINDEX(Ecom_Programs, '" & vPrograms & "') > 0) AND "
      vWhere = vWhere & " (CHARINDEX(Ecom_Source, '" & vSource & "') > 0) "  '...note make this the LAST WHERE else will have training AND

      vSql = vSql & vWhere & " ORDER BY Ecom_CustId, Ecom_Issued, Ecom_CardName "
  
'     sDebug

      sOpenDb
      Set oRs = oDb.Execute(vSql)
      Do While Not oRs.Eof
        sReadEcom     
        sReadCust
  
        '...sometimes accounts have been inactivated or are from a older platform
        If fNoValue(vCust_Title) Then vCust_Title = "(Inactive Account)"      
        If fNoValue(vCust_EcomSplit) Then vCust_EcomSplit = 0
  
        sReadMemb
  
        '...new customer?
        If vEcom_CustId <> vCustIdPrev Then 
  
          If vTotAmount_CA > 0 Then
    %>
    <tr class="h1">
      <th colspan="4">Total :&nbsp;&nbsp;</th>
      <th><%=FormatCurrency(vTotSplitVubz_CA) & "CA"%> </th>
      <th><%=FormatCurrency(vTotSplitCust_CA) & "CA"%> </th>
      <th><%=FormatCurrency(vTotSplitOwnr_CA) & "CA"%> </th>
      <th><%=FormatCurrency(vTotPrice_CA)     & "CA"%> </th>
      <th><%=FormatCurrency(vTotAmount_CA)    & "CA"%> </th>
    </tr>
    <%
          End If
          If vTotAmount_US > 0 Then
    %>
    <tr class="h1">
      <th colspan="4">Total :&nbsp;&nbsp;</th>
      <th><%=FormatCurrency(vTotSplitVubz_US) & "US"%> </th>
      <th><%=FormatCurrency(vTotSplitCust_US) & "US"%> </th>
      <th><%=FormatCurrency(vTotSplitOwnr_US) & "US"%> </th>
      <th><%=FormatCurrency(vTotPrice_US)     & "US"%> </th>
      <th><%=FormatCurrency(vTotAmount_US)    & "US"%> </th>
    </tr>
    <%
          End If
  
          vTotSplitVubz_US  = 0
          vTotSplitVubz_CA  = 0
          vTotSplitCust_US  = 0
          vTotSplitCust_CA  = 0
          vTotSplitOwnr_US  = 0
          vTotSplitOwnr_CA  = 0
          vTotPrice_US      = 0
          vTotPrice_CA      = 0
          vTotAmount_US     = 0
          vTotAmount_CA     = 0
  
        End If
        
        '...get address info
        vAddressInfo = ""
        If vAddress = "Y" Then
          '...only admins can see address cross account
          If svMembLevel = 5 Or Left(svCustId, 4) = Left(vCust_Id, 4) Then
            vAddressInfo = vAddressInfo & "<br>" & vEcom_Address
            vAddressInfo = vAddressInfo & "<br>" & vEcom_City & ", " & vEcom_Province & ", " & vEcom_Country
            vAddressInfo = vAddressInfo & "<br>" & vEcom_Phone
            vAddressInfo = vAddressInfo & "<br>" & vEcom_Email
          End If
        End If
    

        '...if owner (until Apr 2004 we did not carry cardname, so use first/last
        If Len(fOkValue(vEcom_CardName)) = 0 Then vEcom_CardName = vEcom_FirstName & " " & vEcom_LastName

        '...include Order Id, if it exists - Jul 2018
        If Len(vEcom_OrderId) = 0 Then vOrderId = "" Else vOrderId = "&nbsp(" & vEcom_OrderId & ")" End If 

        If svMembLevel < 5 Then
          If Left(svCustId, 4) = Left(vCust_Id, 4) Then
            vNameInfo = "<a " & fStatX & " href='User" & fGroup & ".asp?vMembNo=" & vMemb_No & "'>" & fLeft(vEcom_CardName, 16) & vOrderId & "</a>" & vAddressInfo
          Else
            '...else do not display name info
            vNameInfo = ""
          End If
        Else
          vNameInfo = "<a " & fStatX & " href='EcomEdit.asp?vEcom_No=" & vEcom_No & "'>" & fLeft(vEcom_CardName, 16) & vOrderId & "</a>" & vAddressInfo
        End If
    
        vIssued    = ""
        vExpires   = ""
        vPrice     = ""
        vAmount    = ""
        vSplitVubz = ""
        vSplitCust = ""
        vSplitOwnr = ""
        
        '...get ecom splits
        sGetProg vEcom_Programs '...see if any own splits
    
        If Len(vProg_Owner) > 0 Then
    
          '...owners split (if sales within their channel)
          If Left(vProg_Owner, 4) = Left(vCust_Id, 4) Then
            vEcom_SplitOwnr = vEcom_Prices * vProg_EcomSplitOwner1 / 100
          '...owners split if sales in other channels
          Else
            vEcom_SplitOwnr = vEcom_Prices * vProg_EcomSplitOwner2 / 100
          End If
          
    			'...capture owner totals in table for level 5
    			If svMembLevel = 5 Then
            If Ucase(vEcom_Currency) = "CA" Then
              Redim Preserve aOwnr_CA(vOwnrCnt)
              aOwnr_CA(vOwnrCnt) = aOwnr_CA(vOwnrCnt) + vEcom_SplitOwnr				
            Else
              Redim Preserve aOwnr_US(vOwnrCnt)
              aOwnr_US(vOwnrCnt) = aOwnr_US(vOwnrCnt) + vEcom_SplitOwnr				
            End If
          End If
    
        Else  
          vEcom_SplitOwnr = 0
        End If            
        
        '...compute the channel split from what's left over (unless sold by owner)
        If Left(vProg_Owner, 4) = Left(vCust_Id, 4) Then
          vEcom_SplitCust = 0
        Else
          vEcom_SplitCust = (vEcom_Prices - vEcom_SplitOwnr) * vCust_EcomSplit / 100
        End If
    
        '...vubiz get whats left
        vEcom_SplitVubz = vEcom_Prices - vEcom_SplitCust - vEcom_SplitOwnr
    
    
        '...display the same issue date beside each program
        vIssued = fFormatSqlDate(vEcom_Issued)
            
        '...if vExpires is invalid then get the duration from the customer program string
        On Error Resume Next '...if no customer record, fall thru
        If Not IsDate(vEcom_Expires) Then 
          vExpires = vExpires & fFormatSqlDate(DateAdd("d", fCustProgDuration (vEcom_CustId, vEcom_Programs), vEcom_Issued)) & "&nbsp;"
        Else  
          vExpires = vExpires & fFormatSqlDate(vEcom_Expires) & "&nbsp;"
        End If
        On Error GoTo 0
    
        vPrice     = FormatCurrency(vEcom_Prices,2,0,0) & vEcom_Currency & "&nbsp;"
        vAmount    = FormatCurrency(vEcom_Amount,2,0,0) & vEcom_Currency '...add parm to put a sign in negative nos rather than brackets
        

        '...compute taxes for bottom line of report        
        vTotTAX = vTotTAX + vEcom_Taxes
        If Instr(" NS NB NF ", vEcom_Province) > 0 Then 
          vTotHST = vTotHST + vEcom_Taxes
        ElseIf vEcom_Province = "ON" Then
          If vEcom_Media = "CDs" Then
            vTotPST = vTotPST + vEcom_Taxes * 7/15
            vTotGST = vTotGST + vEcom_Taxes * 8/15
          Else
            vTotGST = vTotGST + vEcom_Taxes
          End If
        ElseIf vEcom_Country = "CA" Then
          vTotGST = vTotGST + vEcom_Taxes
        End If
        
   
        '...get the split values each time a record is read as some use complex forumulae   
        vSplitVubz = FormatCurrency(vEcom_SplitVubz,2,0,0) & vEcom_Currency
        vSplitCust = FormatCurrency(vEcom_SplitCust,2,0,0) & vEcom_Currency
        vSplitOwnr = FormatCurrency(vEcom_SplitOwnr,2,0,0) & vEcom_Currency

        If Ucase(vEcom_Currency) = "CA" Then
          vTotSplitVubz_CA       = vTotSplitVubz_CA       + vEcom_SplitVubz
          vGrandTotSplitVubz_CA  = vGrandTotSplitVubz_CA  + vEcom_SplitVubz
          vTotSplitCust_CA       = vTotSplitCust_CA       + vEcom_SplitCust
          vGrandTotSplitCust_CA  = vGrandTotSplitCust_CA  + vEcom_SplitCust
          vTotSplitOwnr_CA       = vTotSplitOwnr_CA       + vEcom_SplitOwnr
          vGrandTotSplitOwnr_CA  = vGrandTotSplitOwnr_CA  + vEcom_SplitOwnr
          vTotPrice_CA           = vTotPrice_CA           + vEcom_Prices
          vGrandTotPrice_CA      = vGrandTotPrice_CA      + vEcom_Prices
          vTotAmount_CA          = vTotAmount_CA          + vEcom_Amount
          vGrandTotAmount_CA     = vGrandTotAmount_CA     + vEcom_Amount
        Else
          vTotSplitVubz_US       = vTotSplitVubz_US       + vEcom_SplitVubz
          vGrandTotSplitVubz_US  = vGrandTotSplitVubz_US  + vEcom_SplitVubz
          vTotSplitCust_US       = vTotSplitCust_US       + vEcom_SplitCust
          vGrandTotSplitCust_US  = vGrandTotSplitCust_US  + vEcom_SplitCust
          vTotSplitOwnr_US       = vTotSplitOwnr_US       + vEcom_SplitOwnr
          vGrandTotSplitOwnr_US  = vGrandTotSplitOwnr_US  + vEcom_SplitOwnr
          vTotPrice_US           = vTotPrice_US           + vEcom_Prices
          vGrandTotPrice_US      = vGrandTotPrice_US      + vEcom_Prices
          vTotAmount_US          = vTotAmount_US          + vEcom_Amount
          vGrandTotAmount_US     = vGrandTotAmount_US     + vEcom_Amount
        End If

        '... make green if discounts <font color="#008000"> and red for refunds    <font color="#FF0000">
        If vEcom_Amount < 0 Then 
          vPrice     = "<font color='#FF0000'>" & vPrice     & "</font>"
          vAmount    = "<font color='#FF0000'>" & vAmount    & "</font>"
          vSplitVubz = "<font color='#FF0000'>" & vSplitVubz & "</font>"
          vSplitCust = "<font color='#FF0000'>" & vSplitCust & "</font>"
          vSplitOwnr = "<font color='#FF0000'>" & vSplitOwnr & "</font>"

          vTotRefs   = vTotRefs + 1  '...not this is NOT the quantity of programs sold/refunded, but the number of unique programs sold/refunded

        Else
   
          vTotProgs = vTotProgs + 1  '...not this is NOT the quantity of programs sold/refunded, but the number of unique programs sold/refunded

        End If  

        If vEcom_CustId <> vCustIdPrev Then
         
    %>
    <tr>
      <td colspan="9" class="c2" style="padding-top:20px;"><%=vEcom_CustId & "-" & vCust_Title%></td>
    </tr>
    <%
          vCustIdPrev = vEcom_CustId
        End If
       
    %>
    <tr class="d1">
      <td><%=vNameInfo%>&nbsp; </td>
      <td><%=vEcom_Source & " " & fMedia%> <a <%=fStatX%> href="javascript:;" title="<%="P1:" & vProg_EcomSplitOwner1 & " P2:" & vProg_EcomSplitOwner2 & " CH:" & vCust_EcomSplit%>"><%=vEcom_Programs%></a>&nbsp; - <%=fLeft(Trim(fIf(vEcom_Media="Prods", fProdTitle(vEcom_Programs), vProg_Title)), 16)%></td>
      <td><%=vIssued%>&nbsp; </td>
      <td><%=fRemoveBlank(vExpires)%> </td>
      <td><%=fRemoveBlank(vSplitVubz)%> </td>
      <td><%=fRemoveBlank(vSplitCust)%> </td>
      <td><%=fRemoveBlank(vSplitOwnr)%> </td>
      <td><%=fRemoveBlank(vPrice)%> </td>
      <td><%=vAmount%> </td>
    </tr>
    <%

          oRs.MoveNext	        
        Loop
        sCloseDB
    
        '...display totals
        If vTotAmount_CA > 0 Then
    %>
    <tr class="t1">
      <th colspan="4">Total :&nbsp;&nbsp; </th>
      <td style="font-weight: bold; text-align: right;"><%=FormatCurrency(vTotSplitVubz_CA) & "CA"%> </td>
      <th><%=FormatCurrency(vTotSplitCust_CA) & "CA"%> </th>
      <th><%=FormatCurrency(vTotSplitOwnr_CA) & "CA"%> </th>
      <th><%=FormatCurrency(vTotPrice_CA)     & "CA"%> </th>
      <th><%=FormatCurrency(vTotAmount_CA)    & "CA"%> </th>
    </tr>
    <%
        End If
        If vTotAmount_US > 0 Then
    %>
    <tr>
      <th colspan="4">Total :&nbsp;&nbsp; </th>
      <th><%=FormatCurrency(vTotSplitVubz_US) & "US"%> </th>
      <th><%=FormatCurrency(vTotSplitCust_US) & "US"%> </th>
      <th><%=FormatCurrency(vTotSplitOwnr_US) & "US"%> </th>
      <th><%=FormatCurrency(vTotPrice_US)     & "US"%> </th>
      <th><%=FormatCurrency(vTotAmount_US)    & "US"%> </th>
    </tr>
    <%
	      End If
    %>
    <tr>
      <th colspan="9">&nbsp;</th>
    </tr>
    <tr>
      <th colspan="9">&nbsp;</th>
    </tr>
    <%
        If vGrandTotAmount_CA > 0 Then
    %>
    <tr>
      <th class="rowShade">&nbsp;</th>
      <th class="rowShade">&nbsp;</th>
      <th class="rowShade">&nbsp;</th>
      <th class="rowShade">&nbsp;</th>
      <th class="rowShade">$Vubiz&nbsp;&nbsp; </th>
      <th class="rowShade">$Cust&nbsp; </th>
      <th class="rowShade">$Owner&nbsp;&nbsp; </th>
      <th class="rowShade">$Total&nbsp;&nbsp; </th>
      <th class="rowShade">$Total+&nbsp;&nbsp; </th>
    </tr>
    <tr>
      <th colspan="4">Grand Total :&nbsp;&nbsp; </th>
      <th><%=FormatCurrency(vGrandTotSplitVubz_CA) & "CA"%> </th>
      <th><%=FormatCurrency(vGrandTotSplitCust_CA) & "CA"%> </th>
      <th><%=FormatCurrency(vGrandTotSplitOwnr_CA) & "CA"%> </th>
      <th><%=FormatCurrency(vGrandTotPrice_CA)     & "CA"%> </th>
      <th><%=FormatCurrency(vGrandTotAmount_CA)    & "CA"%> </th>
    </tr>
    <%
        End If
        If vGrandTotAmount_US > 0 Then
    %>
    <tr>
      <th colspan="4">Grand Total :&nbsp;&nbsp; </th>
      <th><%=FormatCurrency(vGrandTotSplitVubz_US) & "US"%> </th>
      <th><%=FormatCurrency(vGrandTotSplitCust_US) & "US"%> </th>
      <th><%=FormatCurrency(vGrandTotSplitOwnr_US) & "US"%> </th>
      <th><%=FormatCurrency(vGrandTotPrice_US)     & "US"%> </th>
      <th><%=FormatCurrency(vGrandTotAmount_US)    & "US"%> </th>
    </tr>
    <%
	      End If
    %>
    <tr>
      <th colspan="9">&nbsp;</th>
    </tr>
    <%   If svMembLevel = 5 Then %>
    <tr>
      <th colspan="4" height="22">Owner :&nbsp;&nbsp; </th>
      <th height="22">&nbsp;</th>
      <th height="22">&nbsp;</th>
      <th height="22"><%=FormatCurrency(aOwnr_CA(0)) & "CA"%></th>
      <th colspan="2" height="22">&nbsp; </th>
    </tr>
    <tr>
      <th colspan="4">Owner :&nbsp;&nbsp; </th>
      <th>&nbsp;</th>
      <th>&nbsp;</th>
      <th><%=FormatCurrency(aOwnr_US(0)) & "US"%></th>
      <th colspan="2">&nbsp; </th>
    </tr>
    <tr>
      <th colspan="9">&nbsp;</th>
    </tr>
    <%   End If %>
    <tr>
      <th colspan="9">&nbsp;</th>
    </tr>
    <tr>
      <th colspan="4">Total PST :&nbsp;&nbsp; </th>
      <th>&nbsp;</th>
      <th>&nbsp;</th>
      <th>&nbsp;</th>
      <th colspan="2"><%=FormatCurrency(vTotPST) & "CA"%>&nbsp; </th>
    </tr>
    <tr>
      <th colspan="4">Total GST :&nbsp;&nbsp; </th>
      <th>&nbsp;</th>
      <th>&nbsp;</th>
      <th>&nbsp;</th>
      <th colspan="2"><%=FormatCurrency(vTotGST) & "CA"%>&nbsp; </th>
    </tr>
    <tr>
      <th colspan="4">Total HST :&nbsp;&nbsp; </th>
      <th>&nbsp;</th>
      <th>&nbsp;</th>
      <th>&nbsp;</th>
      <th colspan="2"><%=FormatCurrency(vTotHST) & "CA"%>&nbsp; </th>
    </tr>
    <tr>
      <th colspan="4">Total Tax :&nbsp;&nbsp; </th>
      <th>&nbsp;</th>
      <th>&nbsp;</th>
      <th>&nbsp;</th>
      <th colspan="2"><%=FormatCurrency(vTotTAX) & "CA"%>&nbsp; </th>
    </tr>
    <tr>
      <td colspan="9" style="text-align: center">
        <br /><br /><br />
        <input type="button" onclick="history.back(1)" value="Return" name="bReturn" tabindex="0" class="button">
        <br /><br />
      </td>
    </tr>

    <!--  This ends the section that is created if the form is filled in ok -->
    <% End If %>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


