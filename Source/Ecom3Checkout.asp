<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Routines.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<!--#include virtual = "V5/Inc/EcomCountry.asp"-->
<!--#include virtual = "V5/Inc/Elavon.asp"-->

<html>

<head>
  <title>Ecom3Checkout</title>
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
    '...group checkout

    '...determine media (need to know for PST and shipping reasons)
    vEcom_Media = Session("Ecom_Media")
    If fNoValue(vEcom_Media) Then Response.Redirect "EcomError.asp"

    Session("EcomIssued") = ""
  
    Dim vPrice, vDesc, vQuantity, vTotal, vDiscount, vGst, vPst, vHst, vGstTotal, vPstTotal, vHstTotal, vDuration, vTest, vMsg
    Dim vProdStr, vFlags, vProgFlags, vEcom, vEcomURL, vMerchantNo, vProgram, vPrograms, vLB, vRB
    Dim xxxName, xxxCompany, xxxFirstName, xxxLastName, xxxAddress, xxxCity, xxxPostal, xxxProvince, xxxCountry, xxxPhone, xxxEmail

    '...determine if member uses I/S bypass thus determine Source
    vEcom_Source = "E" '...normal ecom

    If Len(svMembNo) > 0 Then 
      sGetMemb (svMembNo)  
      If vMemb_Ecom Then '...using ecom i/s bypass
        If svMembLevel = 4 Then '...assumes customer
          vEcom_Source = "C"
        Else
          vEcom_Source = "V" '...assumes vubiz
        End If
      End If
    End If   
    
    
      
 
    '...setup order basket array info from Session
    svProdNo = Session("ProdNo")
    If fNoValue(svProdNo) Then svProdNo = 0  
    svProdMax     = Session("ProdMax")
    If fNoValue(svProdMax) Then svProdMax = 0
    If svProdNo   > 0 Then  
      saProd      = Session("Prod")
    Else
      svProdMax   = 0
      Dim saProd()
    End If
  
    vPstTotal      = 0
    vGstTotal      = 0
    vHstTotal      = 0
    vEcom_Amount   = 0
    vEcom_Shipping = 0
  
    '...for submitting braces around ecom flags  
    vLB = Asc("{")
    vRB = Asc("}")
          
    xxxFirstName = Session("xxxFirstName")
    xxxLastName  = Session("xxxLastName")
    xxxName      = xxxFirstName & " " & xxxLastName
    xxxAddress   = Session("xxxAddress")        
    xxxCity      = Session("xxxCity")        
    xxxPostal    = Session("xxxPostal")        
    xxxProvince  = Session("xxxProvince")        
    xxxCountry   = Session("xxxCountry")        
    xxxEmail     = Session("xxxEmail")        
    xxxPhone     = Session("xxxPhone")        
    xxxCompany   = Trim(Session("xxxCompany"))
    
    '...set tax flags
    vFlags = ""
    If fGST(Now, xxxCountry, xxxProvince) Then 
      vFlags = vFlags & Chr(vLB) & "GST" & Chr(vRB)
    End If

    If fHST(Now, xxxCountry, xxxProvince) Then vFlags = vFlags & Chr(vLB) & "HST" & Chr(vRB)
    If fCurrency(xxxCountry) <> "CA" Then vFlags = vFlags & Chr(vLB) & "US" & Chr(vRB)

    '...get values from prod array
    If svProdNo = 0 Then  
  %>
  <table class="table">
    <tr>
      <td style="text-align: center">
        <h6><br><!--[[-->There are no Programs to Checkout.<!--]]--><br></h6>
      </td>
    </tr>
  </table>
  <%
    Else
  %>
  <table class="table">
    <tr>
      <td colspan="3">
        <table class="table">
          <tr>
            <td>
<!--              <script>jTitle("/*--{[--*/Checkout/*--]}--*/", 'Report.jpg')</script>-->
              <img src="../Images/Ecom/Checkout_<%=svLang %>.png" />
            </td>
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

      For i = 1 to svProdMax
        vProgFlags = vFlags '...copy as may get changed if program is tax exempt

        If saProd(2, i) > 0 Then
          vDiscount = saProd(0, i)
          vProgram  = saProd(1, i)
          vQuantity = saProd(2, i)

          vEcom_CatlNo   = vEcom_CatlNo   & Left(vProgram, 8) & "|" 
          vEcom_Programs = vEcom_Programs & Mid(vProgram, 10, 7) & "|" 

          If fCurrency(xxxCountry) <> "CA" Then 
            vEcom_Currency = "US"
            vPrice = saProd(3, i) * vQuantity * (1 - vDiscount/100)
          Else
            vEcom_Currency = "CA"
            vPrice = saProd(4, i) * vQuantity * (1 - vDiscount/100)
            
            '...see if tax exempt then remove from vFlags
            If fProgTaxExempt (vProgram) Then
              vProgFlags = Replace (vFlags,     "{GST}", "")
              vProgFlags = Replace (vProgFlags, "{HST}", "")
              vProgFlags = Replace (vProgFlags, "{PST}", "")
            End If

          End If

          vDuration = Cint(saProd(5, i))
          vDesc = saProd(6, i)

          vEcom_Amount   = vEcom_Amount + vPrice

          vEcom_Quantity = vEcom_Quantity & vQuantity & "|" 
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
      <td><%=Mid(vProgram, 10, 7)%></td>
      <td><%=vDesc%></td>
      <td style="text-align: right"><%=FormatCurrency(vPrice ,2)%></td>
    </tr>
    <%
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
      <td><!--[[-->Tax<!--]]-->@ <%=FormatPercent(fGST(Now, xxxCountry, xxxProvince))%></td>
      <td style="text-align: right"><%=FormatCurrency(vGstTotal ,2)%></td>
    </tr>
    <%
      End If


      If vPstTotal > 0 Then
        vEcom_Amount = vEcom_Amount + vPstTotal
    %>
    <tr>
      <td><!--[[-->PST<!--]]--></td>
      <td><!--[[-->Tax<!--]]-->@ <%=FormatPercent(fPST(Now, xxxCountry, xxxProvince))%></td>
      <td style="text-align: right"><%=FormatCurrency(vPstTotal ,2)%></td>
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
        vEcom_Id = fIf(vMemb_Ecom, "", Session("MembId"))
        vEcom_CustId = Session("CustId")
        vEcom_AcctId = Session("CustAcctId")
        vEcom_Agent = Session("EcomAgent")
        vMemb_FirstName = Session("vMemb_FirstName")
        vMemb_LastName = Session("vMemb_LastName")
        vMemb_Email = Session("vMemb_Email")

        vQuantity = vEcom_Quantity

        Dim vNext   : vNext = "Ecom3GenerateId.asp"
        Dim vSource : vSource = ""

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
  <%
    End If
  %>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>