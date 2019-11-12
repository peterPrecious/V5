<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Routines.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/QueryString.asp"-->

<!--#include virtual = "V5/Inc/EcomCountry.asp"-->
<!--#include virtual = "V5/Inc/Elavon.asp"-->

<html>

<head>
  <title>Ecom6Checkout</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
  Server.Execute vShellHi 
  '...GRTH2884 //vubiz.com/chaccess/GRTH2884/Simulate.asp

  '...do not display the Customer Id and User Id
  Session("Ecom_BypassDisplay") = True

  Dim vCustId, vCustAcctId, vMembFirstName, vMembLastName, vCatl_No, vPrice_CA, vPrice_US
  Dim vPrice, vDesc, vQuantity, vTotal, vDiscount, vGst, vPst, vHst, vGstTotal, vPstTotal, vHstTotal, vDuration, vTest, vMsg
  Dim vFlags, vProgFlags, vEcom, vEcomURL, vMerchantNo, vProgram, vPrograms, vLB, vRB, aProgs, aProg
  Dim xxxName, xxxCompany, xxxFirstName, xxxLastName, xxxAddress, xxxCity, xxxPostal, xxxProvince, xxxCountry, xxxPhone, xxxEmail

  '...store vSource (needed in EcomDisplayId.asp)
  vSource = Request("vSource")        
  sPutQueryString

  vEcom_Media = "Online"
  Session("EcomIssued") = ""
  vEcom_Source = "E" '...normal ecom

  vPstTotal      = 0
  vGstTotal      = 0
  vHstTotal      = 0
  vEcom_Amount   = 0

  '...for submitting braces around ecom flags  
  vLB = Asc("{")
  vRB = Asc("}")
        
  vCustId         = Request("vCustId")        
  vCustAcctId     = Request("vCustAcctId")        

  xxxFirstName    = Request("xxxFirstName")
  xxxLastName     = Request("xxxLastName")
  xxxName         = xxxFirstName & " " & xxxLastName
  xxxCompany      = Trim(Request("xxxCompany"))
  xxxAddress      = Request("xxxAddress")        
  xxxCity         = Request("xxxCity")        
  xxxPostal       = Request("xxxPostal")        
  xxxProvince     = Request("xxxProvince")        
  xxxCountry      = Request("xxxCountry")        
  xxxEmail        = Request("xxxEmail")        
  xxxPhone        = Request("xxxPhone")        

  vEcom_Id        = Request("vEcom_Id")
  vMembFirstName  = fDefault(Request("vMembFirstName"), xxxFirstName)
  vMembLastName   = fDefault(Request("vMembLastName"), xxxLastName)
  vMembEmail      = fDefault(Request("vMembEmail"), xxxEmail)

  '...set tax flags
  vFlags = ""
  If fGST(Now, xxxCountry, xxxProvince) Then 
    vFlags = vFlags & Chr(vLB) & "GST" & Chr(vRB)
  End If

  If fHST(Now, xxxCountry, xxxProvince) Then vFlags = vFlags & Chr(vLB) & "HST" & Chr(vRB)
  If fCurrency(xxxCountry) <> "CA" Then vFlags = vFlags & Chr(vLB) & "US" & Chr(vRB)

  %>

  <table class="table">
     <tr>
      <td colspan="3">
        <table class="table">
          <tr>
            <td><script>jTitle("/*--{[--*/Checkout/*--]}--*/", 'Report.jpg')</script></td>
            <td class="c2">
              <%=xxxName%><br>
              <%=xxxAddress%><br>
              <%=xxxCity & ", " & xxxPostal & ", " & fIf(xxxProvince="None","",xxxProvince & ", ") & xxxCountry%><br><br>
              <%=xxxPhone & " - " & xxxEmail%><br /><br />
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <tr>
      <td class="rowshade" style="text-align: left; width: 20%;"><!--[[-->Program<!--]]--></td>
      <td class="rowshade" style="text-align: left; width: 60%;"><!--[[-->Description<!--]]--></td>
      <td class="rowshade" style="text-align: right; width: 20%;"><!--[[-->Total Price<!--]]-->&nbsp;&nbsp;<%=fCurrency(xxxCountry)%></td>
    </tr>
    <%
      vTotal = Clng(0)

      For Each vFld In Request.Form
        If Left(vFld, 9) = "vProgram_" Then
          vProgFlags = vFlags '...copy as may get changed if program is tax exempt

          aProgs = Split(Request(vFld), "~")
          For i = 0 To Ubound(aProgs)

            aProg = Split(aProgs(i), "|")

            vCatl_No  = aProg(0)
            vProg_Id  = aProg(1)
            vPrice_US = aProg(2)
            vPrice_CA = aProg(3)
            vDesc     = aProg(4)
            vDuration = aProg(5)

            vEcom_CatlNo   = vEcom_CatlNo   & vCatl_No & "|" 
            vEcom_Programs = vEcom_Programs & vProg_Id & "|" 
            
            vQuantity = 1

            If fCurrency(xxxCountry) <> "CA" Then 
              vEcom_Currency = "US"
              vPrice = vPrice_US
            Else
              vEcom_Currency = "CA"
              vPrice = vPrice_CA
              
              '...see if tax exempt then remove from vFlags
              If fProgTaxExempt (vProg_Id) Then
                vProgFlags = Replace (vFlags,     "{GST}", "")
                vProgFlags = Replace (vProgFlags, "{HST}", "")
              End If
            End If

            vEcom_Amount   = vEcom_Amount + vPrice

            vEcom_Quantity = vEcom_Quantity & vQuantity & "|" 
            vEcom_Prices   = vEcom_Prices & vPrice & "|" 
            vEcom_Expires  = vEcom_Expires & fFormatSqlDate(DateAdd("d", vDuration, Now)) & "|" 
  
            '...get detailed taxes for ecom report
            vGst = 0 : vHst = 0
            If Instr(vProgFlags, "GST") > 0 Then vGst = vPrice * fGST(Now, xxxCountry, xxxProvince) 
            If Instr(vProgFlags, "HST") > 0 Then vHst = vPrice * fHST(Now, xxxCountry, xxxProvince)
  
            vEcom_Taxes   = vEcom_Taxes & (vPst + vGst + vHst) & "|" 
  
            '...get detailed taxes for ecom report
            If Instr(vProgFlags, "GST") > 0 Then vGstTotal = vGstTotal + vGst
            If Instr(vProgFlags, "HST") > 0 Then vHstTotal = vHstTotal + vHst
  
    %>
    <tr>
      <td><%=vProg_Id%></td>
      <td><%=vDesc%></td>
      <td style="text-align: right"><%=FormatCurrency(vPrice ,2)%></td>
    </tr>
    <%
          Next
        End If
      Next

      '...remove trailing bars
      vEcom_Quantity = Left(vEcom_Quantity, Len(vEcom_Quantity)-1)
      vEcom_CatlNo   = Left(vEcom_CatlNo,   Len(vEcom_CatlNo)-1)
      vEcom_Programs = Left(vEcom_Programs, Len(vEcom_Programs)-1)
      vEcom_Prices   = Left(vEcom_Prices,   Len(vEcom_Prices)-1)
      vEcom_Taxes    = Left(vEcom_Taxes,    Len(vEcom_Taxes)-1)
      vEcom_Expires  = Left(vEcom_Expires,  Len(vEcom_Expires)-1)

      If vGstTotal > 0 Then
        vEcom_Amount = vEcom_Amount + vGstTotal 
    %>
    <tr>
      <td><!--[[-->GST<!--]]--></td>
      <td><!--[[-->Tax<!--]]--> @ <%=FormatPercent(fGST(Now, xxxCountry, xxxProvince))%></td>
      <td style="text-align: right"><%=FormatCurrency(vGstTotal ,2)%></td>
    </tr>
    <%
        End If
  
  
        If vHstTotal > 0 Then
          vEcom_Amount = vEcom_Amount + vHstTotal 
    %>
    <tr>
      <td><!--[[-->HST<!--]]--></td>
      <td><!--[[-->Tax<!--]]-->&nbsp; @ <%=FormatPercent(fHST(Now, xxxCountry, xxxProvince))%></td>
      <td style="text-align: right"><%=FormatCurrency(vHstTotal ,2)%></td>
    </tr>
    <%
        End If
    %>
    <tr>
      <td class="rowshade" style="text-align: right">&nbsp;</td>
      <td class="rowshade" style="text-align: left"><!--[[-->Total<!--]]--></td>
      <td class="rowshade" style="text-align: right"><%=FormatCurrency(vEcom_Amount,2)%></td>
    </tr>
    <%
      If vEcom_Amount > 0 Then

        '...for consistency       
        vEcom_CustId = vCustId
        vEcom_AcctId = vCustAcctId
        vEcom_Agent = Session("EcomAgent")
        vMemb_FirstName = vMembFirstName
        vMemb_LastName = vMembLastName
        vMemb_Email = vMembEmail

        Dim vNext   : vNext = "Ecom2GenerateId.asp"               
        sPutSqlForm()         '... store values in sql (generate guid and no)
        sSetupElavonForm()   

    %>
    <tr>
      <td colspan="3">
        <form action="<%=vEcomURL%>" method="POST" target="_top">
          <input type="hidden" name="ssl_merchant_id" value="<%=ssl_merchant_id%>">
          <input type="hidden" name="ssl_user_id" value="<%=ssl_user_id%>">
          <input type="hidden" name="ssl_pin" value="<%=ssl_pin%>">
          <input type="hidden" name="ssl_transaction_type" value="ccsale">
          <input type="hidden" name="ssl_show_form" value="true">
          <input type="hidden" name="ssl_test_mode" value="<%=ssl_test_mode%>">
          <input type="hidden" name="ssl_amount" value="<%=ssl_amount%>">
          <input type="hidden" name="ssl_invoice_number" value="<%=ssl_invoice_number%>">

          <input type="hidden" name="ssl_company" value="<%=ssl_company%>">
          <input type="hidden" name="ssl_first_name" value="<%=ssl_first_name%>">
          <input type="hidden" name="ssl_last_name" value="<%=ssl_last_name%>">
          <input type="hidden" name="ssl_avs_address" value="<%=ssl_avs_address%>">
          <input type="hidden" name="ssl_city" value="<%=ssl_city%>">
          <input type="hidden" name="ssl_state" value="<%=ssl_state%>">
          <input type="hidden" name="ssl_avs_zip" value="<%=ssl_avs_zip%>">
          <input type="hidden" name="ssl_country" value="<%=ssl_country%>">
          <input type="hidden" name="ssl_email" value="<%=ssl_email%>">

          <input type="hidden" name="ssl_result_format" value="HTML">
          <input type="hidden" name="ssl_receipt_link_method" value="POST">
          <input type="hidden" name="ssl_receipt_link_url" value="<%=ssl_receipt_link_url%>">
          <input type="hidden" name="order_guid" value="<%=order_guid%>">

          <table class="table">
            <tr>
              <td style="text-align: center">
                <h6><br /><!--[[-->IMPORTANT<!--]]-->...</h6>
                <h6 style="text-align: left">
                  <!--[[-->Clicking <b>Continue</b> below will transfer you to &quot;Elavon&quot; where you can make a secure payment of<!--]]-->&nbsp;<%=FormatCurrency(vEcom_Amount,2) & fCurrency(xxxCountry)%>.
                  <!--[[-->This order will appear on your credit card statement as VUBIZ.COM LTD.<!--]]-->
                  <!--[[-->After payment is approved, you MUST click on the &quot;Finish Transaction and Return to Store&quot; button so we can setup the Programs you just ordered.<!--]]-->
                </h6>
                <h6><!--[[-->ONLY CLICK <b>Continue</b> ONCE!<!--]]--></h6>
                <p>
                  <input type="button" onclick="history.back(1)" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button"><%=f10%>
                  <input type="submit" value="<%=bContinue%>" name="bContinue" class="button">
                </p>
              </td>
            </tr>

            <tr>
              <td style="text-align: center">
                <br />
                <img border="0" src="../Images/Common/Visa.gif" width="37" height="23">
                <img border="0" src="../Images/Common/Mastercard.gif" width="37" height="23">
                <img border="0" src="../Images/Common/Amex.gif" width="37" height="23">
              </td>
            </tr>

          </table>

        </form>
      </td>
    </tr>
    <%
      End If
    %>  
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>