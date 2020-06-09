
<% 
  Dim ssl_merchant_id, ssl_user_id, ssl_pin, ssl_test_mode, ssl_transaction_currency, ssl_receipt_link_url, ssl_amount, ssl_invoice_number, ssl_company, ssl_first_name, ssl_last_name, ssl_avs_address, ssl_city, ssl_state, ssl_avs_zip, ssl_country, ssl_phone, ssl_email
  Dim order_no, order_guid


  Sub sPutSqlForm() '...store ecom values
    sOpenCmdApp
    With oCmdApp
      .CommandText = "sp5elavonPutLogs"
  
      .Parameters.Append .CreateParameter("@xxxName", adVarChar, adParamInput, 50, xxxName)
      .Parameters.Append .CreateParameter("@xxxFirstName", adVarChar, adParamInput, 50, xxxFirstName)
      .Parameters.Append .CreateParameter("@xxxLastName", adVarChar, adParamInput, 50, xxxLastName)
      .Parameters.Append .CreateParameter("@xxxEmail", adVarChar, adParamInput, 50, xxxEmail)
      .Parameters.Append .CreateParameter("@xxxCompany", adVarChar, adParamInput, 250, xxxCompany)
      .Parameters.Append .CreateParameter("@xxxAddress", adVarChar, adParamInput, 50, xxxAddress)
      .Parameters.Append .CreateParameter("@xxxCity", adVarChar, adParamInput, 50, xxxCity)
      .Parameters.Append .CreateParameter("@xxxPostal", adVarChar, adParamInput, 50, xxxPostal)
      .Parameters.Append .CreateParameter("@xxxProvince", adVarChar, adParamInput, 50, xxxProvince)
      .Parameters.Append .CreateParameter("@xxxCountry", adVarChar, adParamInput, 50, xxxCountry)
      .Parameters.Append .CreateParameter("@xxxPhone", adVarChar, adParamInput, 50, xxxPhone)

      .Parameters.Append .CreateParameter("@vMemb_FirstName", adVarChar, adParamInput, 50, vMemb_FirstName)
      .Parameters.Append .CreateParameter("@vMemb_LastName", adVarChar, adParamInput, 50, vMemb_LastName)
      .Parameters.Append .CreateParameter("@vMemb_Email", adVarChar, adParamInput, 50, vMemb_Email)

      .Parameters.Append .CreateParameter("@vEcom_Id", adVarChar, adParamInput, 50, vEcom_Id)
      .Parameters.Append .CreateParameter("@vEcom_CustId", adVarChar, adParamInput, 50, vEcom_CustId)
      .Parameters.Append .CreateParameter("@vEcom_AcctId", adVarChar, adParamInput, 50, vEcom_AcctId)
      .Parameters.Append .CreateParameter("@vEcom_Agent", adVarChar, adParamInput, 50, vEcom_Agent)

      .Parameters.Append .CreateParameter("@vEcom_CatlNo", adVarChar, adParamInput, 4000, Trim(vEcom_CatlNo))
      .Parameters.Append .CreateParameter("@vEcom_Programs", adVarChar, adParamInput, 4000, Trim(vEcom_Programs))
      .Parameters.Append .CreateParameter("@vEcom_Quantity", adVarChar, adParamInput, 4000, Trim(vEcom_Quantity))
      .Parameters.Append .CreateParameter("@vEcom_Prices", adVarChar, adParamInput, 4000, Trim(vEcom_Prices))
      .Parameters.Append .CreateParameter("@vEcom_Taxes", adVarChar, adParamInput, 4000, Trim(vEcom_Taxes))
      .Parameters.Append .CreateParameter("@vEcom_Expires", adVarChar, adParamInput, 4000, Trim(vEcom_Expires))

      .Parameters.Append .CreateParameter("@vEcom_Amount", adVarChar, adParamInput, 50, FormatNumber(vEcom_Amount, 2,,,0))
      .Parameters.Append .CreateParameter("@vEcom_Lang", adVarChar, adParamInput, 50, vEcom_Lang)
      .Parameters.Append .CreateParameter("@vEcom_Currency", adVarChar, adParamInput, 50, vEcom_Currency)
      .Parameters.Append .CreateParameter("@vEcom_Media", adVarChar, adParamInput, 50, vEcom_Media)
      .Parameters.Append .CreateParameter("@vEcom_OrderNo", adVarChar, adParamInput, 50, vEcom_OrderNo)
      .Parameters.Append .CreateParameter("@vEcom_Source", adVarChar, adParamInput, 50, vEcom_Source)

      .Parameters.Append .CreateParameter("@vNext", adVarChar, adParamInput, 50, vNext)
      .Parameters.Append .CreateParameter("@vSource", adVarChar, adParamInput, 250, vSource)
    End With

	  Set oRs = oCmdApp.Execute()
    If Not oRs.Eof Then 
      order_no = oRs("no")
      order_guid = Replace(Replace(oRs("guid"), "{", ""), "}", "")    '...strip off braces around guid returned from SQL
    End If        
    Set oRs = Nothing
    Set oCmdApp = Nothing
    sCloseDbApp

  End Sub




  Sub sSetupElavonForm()    '... populate elavon form

'    If vTest = "Y" Then
'      vEcomURL   = "EcomPatience.asp"
'      vMsg       = "<!--{{-->We are now completing your ecommerce transaction.<!--}}-->"      '...This error message is for the "Patience" screen 

    If Lcase(xxxEmail) = "pbulloch@vubiz.com" Or vMemb_Ecom Or svEcomBypass Then
'     vTest      = "y"
'     vMsg       = "<!--{{-->We are now completing your ecommerce transaction.<!--}}-->"      '...This error message is for the "Patience" screen 
'     vEcomURL   = "EcomPatience.asp"  
      vEcomURL   = vNext
      Session("BypassEcom") = order_no
    Else

'     If svServer = "stagingweb.vubiz.com" Or svServer = "corporate.vubiz.com"Then
        vEcomURL   = "https://api.convergepay.com/VirtualMerchant/process.do"                 '...new SHA 2 / TLS 1.2 service
'     Else
'       vEcomURL   = "https://www.myvirtualmerchant.com/VirtualMerchant/process.do"           '...legacy service
'     End If

'     vMsg       = "" 
      Session("BypassEcom") = ""
    End If

 
    '...live parameters
    ssl_merchant_id = "679356" 
    ssl_user_id = "webpage"
    ssl_test_mode = false

    If vEcom_Currency = "US" Then
      ssl_transaction_currency = "USD"
      ssl_pin = "3554"
    Else
      ssl_transaction_currency = "CAD"
      ssl_pin = "8404"
    End If

    '...demo (put block before next block to test)
'   ssl_merchant_id = "001474" 
'   ssl_user_id = "webpage"
'   ssl_pin = "0UCCYH"
'   ssl_test_mode = true

'   vEcomURL   = "https://api.demo.convergepay.com/VirtualMerchantDemo/process.do" '...test new SHA2 and TLS 1.2
'   ssl_merchant_id = "000127" 
'   ssl_user_id = "ssltest"
'   ssl_pin = "IERAOBEE5V0D6Q3Q6R51TG89XAIVGEQ3LGLKMKCKCVQBGGGAU7FN627GPA54P5HR"
'   ssl_test_mode = true
  
    ssl_amount = vEcom_Amount
    ssl_invoice_number = order_no
    ssl_company = xxxCompany
    ssl_first_name = xxxFirstName
    ssl_last_name = xxxLastName
    ssl_avs_address = xxxAddress
    ssl_city = xxxCity
    ssl_state = xxxProvince
    ssl_avs_zip = xxxPostal
    ssl_country = sp5countryCodeElavon (xxxCountry)
    ssl_phone = xxxPhone
    ssl_email = xxxEmail

'   ...this does not handle SSL nor does adding "//" in front - need full URL below
'   ssl_receipt_link_url = svServer & "/V5/Source/" & vNext
'   ssl_receipt_link_url = svServer & "/V5/Code/" & vNext

   If (svSSL) Then
     ssl_receipt_link_url = "https://" + svServer & "/V5/Code/" & vNext
   Else
     ssl_receipt_link_url = svServer & "/V5/Code/" & vNext
   End If

  End Sub


  Sub sGetSqlForm(orderNo) '...get ecom values
  
    vEcom_Eof = True

    sOpenCmdApp
    With oCmdApp
      .CommandText = "sp5elavonGetLogs" 
      .Parameters.Append .CreateParameter("@no", adInteger, adParamInput, , orderNo)
    End With
	  Set oRs = oCmdApp.Execute()
    If Not oRs.Eof Then 

      vEcom_Eof = False
      vSource = oRs("vSource")

      vEcom_CustId = oRs("vEcom_CustId")
      vEcom_AcctId = oRs("vEcom_AcctId")
      vEcom_Agent = oRs("vEcom_Agent")
      vEcom_Id = oRs("vEcom_Id")
      vEcom_CatlNo = oRs("vEcom_CatlNo")
      vEcom_Programs = oRs("vEcom_Programs")
      vEcom_Prices = oRs("vEcom_Prices")
      vEcom_Taxes = oRs("vEcom_Taxes")
      vEcom_Expires = oRs("vEcom_Expires")
      vEcom_Amount = oRs("vEcom_Amount")
      vEcom_Currency = oRs("vEcom_Currency")
      vEcom_Lang = oRs("vEcom_Lang")
      vEcom_Quantity = oRs("vEcom_Quantity")
      vEcom_Media = oRs("vEcom_Media")
      vEcom_Source = oRs("vEcom_Source")

      vEcom_CardName = oRs("xxxName")
      vEcom_Address = oRs("xxxAddress")
      vEcom_City = oRs("xxxCity")
      vEcom_Postal = oRs("xxxPostal")
      vEcom_Province = oRs("xxxProvince")
      vEcom_Country = oRs("xxxCountry")
      vEcom_Phone = oRs("xxxPhone")

      vEcom_Organization = oRs("xxxCompany")      
      vMemb_Organization = oRs("xxxCompany")      

      vMemb_FirstName = oRs("vMemb_FirstName")
      vMemb_LastName = oRs("vMemb_LastName")
      vMemb_Email = oRs("vMemb_Email")

      vEcom_FirstName      = vMemb_FirstName
      vEcom_LastName       = vMemb_LastName
      vEcom_Email          = vMemb_Email  

    End If
    Set oRs = Nothing
    Set oCmdApp = Nothing
    sCloseDbApp


    vEcom_Shipping  = 0  '...this was removed from the SQL

  End Sub



  '...these are used by the utility ElavonPost.asp to manually post a transaction when the customer forgot to return to Vubiz
  
  Function sp5elavonToEcom(orderNo) '...check if this order go to Vubiz, it will return the number of records with this orderNo
    sOpenCmdApp
    With oCmdApp
      .CommandText = "sp5elavonToEcom" 
      .Parameters.Append .CreateParameter("@orderNo", adInteger, adParamInput, , orderNo)
    End With
	  Set oRs = oCmdApp.Execute()
    sp5ElavonToEcom = oRs("Count")
    Set oRs = Nothing
    Set oCmdApp = Nothing
    sCloseDbApp
  End Function



%>