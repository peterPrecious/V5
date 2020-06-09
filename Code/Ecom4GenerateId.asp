<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Routines.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Debug_Routines.asp"-->

<% 
  Dim vTest, aPrograms, aPrices, aTaxes, vExpires, aExpires, aAmounts, vTotalPrices, vTotalAmount

  '...determine if test or live
  vTest = fDefault(Request("vEcom_Test"), "y")

  '...get querystring order values from either InternetSecure (test=n) or Ecom2Checkout (test=y)
  '   build up values for email tracking/debugging (stripping out any br tags)
  For Each vFld In Request.Form
    vValue = fUnQuote(Request.Form(vFld))

    Select Case vFld
      Case "vEcom_CustId"         : vEcom_CustId         = vValue
      Case "vEcom_AcctId"         : vEcom_AcctId         = vValue
      Case "vEcom_Agent"          : vEcom_Agent          = vValue
      Case "vEcom_Id"             : vEcom_Id             = vValue
      Case "vEcom_Programs"       : vEcom_Programs       = vValue
      Case "vEcom_Prices"         : vEcom_Prices         = vValue
      Case "vEcom_Taxes"          : vEcom_Taxes          = vValue
      Case "vEcom_Expires"        : vEcom_Expires        = vValue
      Case "vEcom_Amount"         : vEcom_Amount         = vValue
      Case "vEcom_Currency"       : vEcom_Currency       = vValue
      Case "vEcom_Lang"           : vEcom_Lang           = vValue
      Case "vEcom_Quantity"       : vEcom_Quantity       = vValue
      Case "vEcom_Media"          : vEcom_Media          = vValue
      Case "vEcom_OrderNo"        : vEcom_OrderNo        = vValue
      Case "vEcom_Shipping"       : vEcom_Shipping       = vValue
      Case "vEcom_Source"         : vEcom_Source         = vValue
      Case "receiptnumber"        : vEcom_InternetSecure = vValue
      Case "xxxName"              : vEcom_CardName       = vValue
      Case "xxxAddress"           : vEcom_Address        = vValue
      Case "xxxCity"              : vEcom_City           = vValue
      Case "xxxPostal"            : vEcom_Postal         = vValue
      Case "xxxProvince"          : vEcom_Province       = vValue
      Case "xxxCountry"           : vEcom_Country        = vValue
      Case "xxxPhone"             : vEcom_Phone          = vValue
      Case "vMemb_FirstName"      : vMemb_FirstName      = vValue : vEcom_FirstName      = vValue
      Case "vMemb_LastName"       : vMemb_LastName       = vValue : vEcom_LastName       = vValue
      Case "vMemb_Email"          : vMemb_Email          = vValue : vEcom_Email          = vValue
      Case "xxxCompany"
        '...strip off the i/s company notice 
        i = Instr(vValue, "(")
        If Len(vValue) = 0 Or i = 1 Then
            												vMemb_Organization   = ""
            												vEcom_Organization   = ""
        ElseIf i > 1 Then
            												vMemb_Organization   = Trim(Left(vValue, i-1))
            												vEcom_Organization   = vMemb_Organization
        Else
            												vMemb_Organization   = vValue
            												vEcom_Organization   = vValue
        End If
      End Select    
    End Select
  Next  
  

  '...double issue?
  If fIsNewEcom = False Then 
    Response.Redirect "EcomError.asp?vMsg=" & Server.UrlEncode("This is a duplicate transaction.")   
  '...invalid response from InternetSecure or user screwing around?
  ElseIf Session("EcomIssued") = True Or Len(vEcom_Amount) = 0 Or Len(vEcom_Programs) = 0 Or Len(vEcom_CustId) = 0 Or Len(vEcom_Media) = 0 Then 
    Response.Redirect "EcomError.asp?vMsg=" & Server.UrlEncode("This is an invalid transaction.")  
  End If

  vMemb_AcctId        = vEcom_AcctId
  vMemb_No            = fNextMembNo (vEcom_AcctId)

  vMemb_Level         = 2
  vMemb_Expires       = vEcom_Expires
  vMemb_Memo          = "Ecom: " & vEcom_OrderNo
  sAddMemb vMemb_AcctId

  vEcom_Issued        = fFormatSqlDate (Now)
  vEcom_Shipping      = 0
  vEcom_Lang          = svLang
  vEcom_Quantity      = 1
  vEcom_MembNo        = vMemb_No
  vEcom_Id            = vMemb_Id

  sAddEcom vMemb_AcctId


  '...Store in Session variables so not visible on url
  Session("EcomCust")   = vEcom_CustId
  Session("EcomId")     = vEcom_Id
  Session("EcomIssued") = True

  
  '...display details to the customer
  Response.Redirect "Ecom4DisplayId.asp?vClose=Y"
 
  '...Generate the security code (also in Db_Memb.asp)... not used?
  Function fSecurityCodex (vCustNo, vMembNo)
    Dim vTemp
    Const cAlpha = "ABCDEFGHXY"
    vTemp = vMembNo * 4141
    vTemp = vMembNo * 4141 + Right(vCustNo, 4)
    vTemp = Right("0000" & vTemp, 4)
    fSecurityCode = ""
    For i = 1 To 4
      j = Mid(vTemp,i,1)   
      k = Mid(cAlpha, j+1, 1)
      fSecurityCode = fSecurityCode & k
    Next
  End Function

%>

