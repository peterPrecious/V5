<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Routines.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<html>

<head>
  <title>Ecom0Checkout</title>
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
</head>

<body>

  <%     
    Server.Execute vShellHi 

    Dim vPrice, vDesc, vQuantity, vTotal, vDiscount, vGst, vPst, vHst, vGstTotal, vPstTotal, vHstTotal, vDuration, vQuantityShipping, vTest, vMsg
    Dim vProdStr, vFlags, vProgFlags, vEcom, vEcomURL, vMerchantNo, vProgram, vPrograms, vLB, vRB
    Dim xxxName, xxxCompany, xxxFirstName, xxxLastName, xxxAddress, xxxCity, xxxPostal, xxxProvince, xxxCountry, xxxPhone, xxxEmail

    Dim ssl_merchant_id,ssl_user_id,ssl_pin,ssl_test_mode,ssl_description,ssl_transaction_currency,ssl_receipt_link_url,ssl_receipt_link_text, ssl_company,ssl_first_name,ssl_last_name,ssl_city,ssl_state,ssl_country,ssl_email


    vEcom_Amount = "1.00"

    '...live
    ssl_merchant_id = "679356" 
    ssl_user_id = "webpage"
    ssl_pin = "4055"

    '...demo
    ssl_merchant_id = "001474" 
    ssl_user_id = "webpage"
    ssl_pin = "0UCCYH"


    ssl_test_mode = True
    ssl_description = "<!--{{-->eLearning Course<!--}}-->"

    If vEcom_Currency = "US" Then
      ssl_transaction_currency = "USD"
      ssl_pin = "4055"
    Else
      ssl_transaction_currency = "CAD"
      ssl_pin = "8404"
      ssl_pin = "0UCCYH" '...demo
    End If


    ssl_receipt_link_url = "//stagingweb.vubiz.com/v6/ecomGenerateId.aspx"

        '...if testing or special member (ie vMemb_Ecom = True) then bypass InternetSecure 
    If Lcase(xxxEmail) = "pbulloch@vubiz.com" Or vMemb_Ecom Or svEcomBypass Then
      vTest      = "y"
      vEcomURL   = "EcomPatience.asp"
    Else
      vTest      = "n"
      vEcomURL   = "https://www.myvirtualmerchant.com/VirtualMerchant/process.do"
      vEcomURL   = "https://demo.myvirtualmerchant.com/VirtualMerchantDemo/process.do"
    End If


      
        '...create timestamp orderno
    vEcom_Orderno = Right("00" & Year(Now), 2) & Right("00" & Month(Now), 2) & "-" & Right("00" & Day(Now), 2) & Right("00" & Hour(Now), 2) & "-" & Right("00" & Minute(Now), 2) & Right("00" & Second(Now), 2) 



  %>


  <table class="table">
    <tr>
      <td>Elavon Tester</td>
    </tr>

    <tr>
      <td>
        <form action="<%=vEcomURL%>" method="POST" target="_top">

          <!-- Elavon Values -->

          <!-- create these next 3 field on submit so they are not visible to client -->
          <input type="hidden" name="ssl_merchant_id" value="<%=ssl_merchant_id%>">
          <input type="hidden" name="ssl_user_id" value="<%=ssl_user_id%>">
          <input type="hidden" name="ssl_pin" value="<%=ssl_pin%>">

          <input type="hidden" name="ssl_transaction_type" value="ccsale">
          <input type="hidden" name="ssl_show_form" value="true">
          <input type="hidden" name="ssl_test_mode" value="<%=ssl_test_mode%>">
          <input type="hidden" name="ssl_amount" value="<%=vEcom_Amount%>">
          <input type="hidden" name="ssl_invoice_number" value="<%=vEcom_Orderno%>">
          <input type="hidden" name="ssl_description" value="<%=ssl_description%>">

          <input type="hidden" name="ssl_result_format" value="HTML">
          <input type="hidden" name="ssl_receipt_link_method" value="POST">
          <input type="hidden" name="ssl_receipt_link_url" value="<%=ssl_receipt_link_url%>">
          <input type="hidden" name="ssl_receipt_link_text" value="<%=ssl_receipt_link_text%>">

          <input type="hidden" name="ssl_avs_address" value="<%=xxxAddress%>">
          <input type="hidden" name="ssl_avs_zip" value="<%=xxxPostal%>">

          <input type="submit" onclick="this.disabled = true" value="<%=bContinue%>" name="bContinue" class="button">
        </form>
      </td>
    </tr>

  </table>


</body>
</html>

