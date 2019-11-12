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
  <title>Ecom5Checkout</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>AC_FL_RunContent = 0;</script>
  <script src="/V5/Inc/AC_RunActiveContent.js"></script>
  <script>
    function jTitle (vTitle, vImage) {
      var vParm = "title=" + vTitle + '&image=/V5/Images/Titles/' + vImage;
      AC_FL_RunContent('codebase','//download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0','name','flashVars','width','265','height','85','align','middle','id','flashVars','src','/V5/Images/Titles/VuTitles','FlashVars',vParm,'quality','high','bgcolor','#ffffff','allowscriptaccess','sameDomain','allowfullscreen','false','pluginspage','///go/getflashplayer','movie','/V5/Images/Titles/VuTitles');
    }
  </script>
</head>

<body>

  <% 
  Server.Execute vShellHi 
  '...chaccess/hospital/hecs //vubiz.com/chaccess/hospital/hecs

  vEcom_Media = "Online"
  Session("EcomIssued") = ""

  '...store vSource
  vSource = Request("vSource")        
  sPutQueryString

  Dim vPrice, vDesc, vQuantity, vDiscount, vGst, vPst, vHst, vGstTotal, vPstTotal, vHstTotal, vDuration, vTest, vMsg
  Dim vFlags, vProgFlags, vEcom, vEcomURL, vMerchantNo, vProgram, vPrograms, vLB, vRB
  Dim xxxName, xxxCompany, xxxFirstName, xxxLastName, xxxAddress, xxxCity, xxxPostal, xxxProvince, xxxCountry, xxxPhone, xxxEmail
  Dim vCustId, vCustAcctId, vMembFirstName, vMembLastName

  '...determine if member uses I/S bypass thus determine Source
  vEcom_Source   = "E" '...normal ecom

  vPstTotal      = 0
  vGstTotal      = 0
  vHstTotal      = 0
  vEcom_Amount   = 0

  '...for submitting braces around ecom flags  
  vLB = Asc("{")
  vRB = Asc("}")

  vPrograms       = Request("vPrograms")        
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
            <td><script>jTitle("<%=fPhraH(000916)%>", 'Report.jpg')</script></td>
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
      <td class="rowshade" style="text-align: left; width: 20%;"><!--webbot bot='PurpleText' PREVIEW='Program'--><%=fPhra(000201)%></td>
      <td class="rowshade" style="text-align: left; width: 60%;"><!--webbot bot='PurpleText' PREVIEW='Description'--><%=fPhra(000118)%></td>
      <td class="rowshade" style="text-align: right; width: 20%;"><!--webbot bot='PurpleText' PREVIEW='Total Price'--><%=fPhra(000387)%>&nbsp;&nbsp;<%=fCurrency(xxxCountry)%></td>
    </tr>
    <%
      '...assume just one program coming through (else need to create a loop)
      '...format:  "90.00|100.00|P1269EN|Online Preparation for Childbirth|90"

      vProgFlags = vFlags '...copy as may get changed if program is tax exempt

      Dim aProg
      aProg = Split(vPrograms, "|")

      If fCurrency(xxxCountry) <> "CA" Then 
        vEcom_Currency = "US"
        vPrice = aProg (0)
      Else
        vEcom_Currency = "CA"
        vPrice = aProg (1)
        
        '...see if tax exempt then remove from vFlags
        If fProgTaxExempt (vEcom_Programs) Then
          vFlags = Replace (vFlags, "{GST}", "")
          vFlags = Replace (vFlags, "{HST}", "")
          vFlags = Replace (vFlags, "{PST}", "")
        End If

      End If

      vProgram = aProg(2)
      vEcom_Programs = aProg(2) & "|"
      vDesc = aProg(3)
      vDuration = aProg(4)

      vEcom_Amount   = vEcom_Amount + vPrice

      vEcom_Quantity = vEcom_Quantity & 1 & "|" 
      vEcom_Prices   = vEcom_Prices & vPrice & "|" 
      vEcom_Expires  = vEcom_Expires & fFormatSqlDate(DateAdd("d", vDuration, Now)) & "|" 

      '...get detailed taxes for ecom report
      vPst = 0 : vGst = 0 : vHst = 0
      If Instr(vProgFlags, "PST") > 0 Then vPst = vPrice * fPST(Now, xxxCountry, xxxProvince)
      If Instr(vProgFlags, "GST") > 0 Then vGst = vPrice * fGST(Now, xxxCountry, xxxProvince) 
      If Instr(vProgFlags, "HST") > 0 Then vHst = vPrice * fHST(Now, xxxCountry, xxxProvince)

      vEcom_Taxes   = vEcom_Taxes & (vPst + vGst + vHst) & "|" 

      '...get detailed taxes for ecom report
      If Instr(vProgFlags, "PST") > 0 Then vPstTotal = vPstTotal + vPst
      If Instr(vProgFlags, "GST") > 0 Then vGstTotal = vGstTotal + vGst
      If Instr(vProgFlags, "HST") > 0 Then vHstTotal = vHstTotal + vHst

    %>
    <tr>
      <td><%=vProgram%></td>
      <td><%=vDesc%></td>
      <td style="text-align: right"><%=FormatCurrency(vPrice ,2)%></td>
    </tr>
    <%

      vEcom_Quantity = Left(vEcom_Quantity, Len(vEcom_Quantity)-1)
      vEcom_Programs = Left(vEcom_Programs, Len(vEcom_Programs)-1)
      vEcom_Prices   = Left(vEcom_Prices,   Len(vEcom_Prices)-1)
      vEcom_Taxes    = Left(vEcom_Taxes,    Len(vEcom_Taxes)-1)
      vEcom_Expires  = Left(vEcom_Expires,  Len(vEcom_Expires)-1)


      If vGstTotal > 0 Then
        vEcom_Amount = vEcom_Amount + vGstTotal 
    %>
    <tr>
      <td>GST</td>
      <td>Tax @ <%=FormatPercent(fGST(Now, xxxCountry, xxxProvince))%></td>
      <td style="text-align: right"><%=FormatCurrency(vGstTotal ,2)%></td>
    </tr>
    <%
      End If

      If vPstTotal > 0 Then
        vEcom_Amount = vEcom_Amount + vPstTotal
    %>
    <tr>
      <td>PST</td>
      <td>Tax @ <%=FormatPercent(fPST(Now, xxxCountry, xxxProvince))%></td>
      <td style="text-align: right"><%=FormatCurrency(vPstTotal ,2)%></td>
    </tr>
    <%
      End If

      If vHstTotal > 0 Then
        vEcom_Amount = vEcom_Amount + vHstTotal 
    %>
    <tr>
      <td>HST</td>
      <td>Tax @ <%=FormatPercent(fHST(Now, xxxCountry, xxxProvince))%></td>
      <td style="text-align: right"><%=FormatCurrency(vHstTotal ,2)%></td>
    </tr>
    <%
      End If
    %>
    <tr>
      <td class="rowshade" style="text-align: right">&nbsp;</td>
      <td class="rowshade" style="text-align: left"><!--webbot bot='PurpleText' PREVIEW='Total'--><%=fPhra(000020)%></td>
      <td class="rowshade" style="text-align: right"><%=FormatCurrency(vEcom_Amount,2)%></td>
    </tr>
    <%
      If vEcom_Amount > 0 Then

        '...for consistency  
        vEcom_Id = ""
        vEcom_CustId = vCustId
        vEcom_AcctId = vCustAcctId
        vEcom_Agent = ""
        vMemb_FirstName = vMembFirstName
        vMemb_LastName = vMembLastName
        vMemb_Email = vMembEmail

        vEcom_CatlNo = 0

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
                <h6><br /><!--webbot bot='PurpleText' PREVIEW='IMPORTANT'--><%=fPhra(000153)%>...</h6>
                <h6 style="text-align: left">
                  <!--webbot bot='PurpleText' PREVIEW='Clicking <b>Continue</b> below will transfer you to &quot;Elavon&quot; where you can make a secure payment of'--><%=fPhra(001777)%>&nbsp;<%=FormatCurrency(vEcom_Amount,2) & fCurrency(xxxCountry)%>.
                  <!--webbot bot='PurpleText' PREVIEW='This order will appear on your credit card statement as VUBIZ.COM LTD.'--><%=fPhra(000925)%>
                  <!--webbot bot='PurpleText' PREVIEW='After payment is approved, you MUST click on the &quot;Finish Transaction and Return to Store&quot; button so we can setup the Programs you just ordered.'--><%=fPhra(001778)%>
                </h6>
                <h6><!--webbot bot='PurpleText' PREVIEW='ONLY CLICK <b>Continue</b> ONCE!'--><%=fPhra(000315)%></h6>
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

