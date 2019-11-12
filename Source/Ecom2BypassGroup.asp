<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/QueryString.asp"-->

<%

  '...This allows customers to bypass the selection and basket and go straight to the customer input screen for single user license.
  '   The programs selected for the basket are passed into this page via the vMemo field 
  '   which must come in via the original launch string and passed through via the querystring

  '...Note, the programs must be exactly as they appear on the normal ecom content page, separated by a pipe
  '   the last value is now they must arrive
  '   use this example to build the url which looks like this:
  '   /V5/default.asp?vCust=1234EN&vAction=qOrderSingle&vMemo=P1014EN%7E79%7E125%7E79%7E125%7E90%7EFeelings%3A+Quality+Service%2E%2E%2EFirst+Time%2C+Every+Time+for+Professionals%7CP1017EN%7E79%7E125%7E79%7E125%7E90%7EFeelings%3A+Customer+Service+for+those+in+Service+Retail

  ' vMemo = "P1014EN~79~125~79~125~90~Feelings: Quality Service...First Time, Every Time for Professionals|" _
  '       & "P1017EN~79~125~79~125~90~Feelings: Customer Service for those in Service Retail"
  ' vMemo = Server.UrlEncode(vMemo)
  ' vMemo ="P1014EN%7E79%7E125%7E79%7E125%7E90%7EFeelings%3A+Quality+Service%2E%2E%2EFirst+Time%2C+Every+Time+for+Professionals%7CP1017EN%7E79%7E125%7E79%7E125%7E90%7EFeelings%3A+Customer+Service+for+those+in+Service+Retail"

 
  Dim aPrograms, aValue, svProd_No, svProd_Max, vEcom_Quantity
  
  Session("Ecom_Media") = "Group"
  
  sGetQueryString  '...this gets the vMemo field (with all the other fields)
  aPrograms = Split(vMemo, "|")

  '...put values in prod array
  svProd_No  = 0
  svProd_Max = 0

  '...put all the selected items into an array
  For i = 0 to Ubound(aPrograms)
  
    vValue                = Replace(aPrograms(i), "'", " ")
    aValue                = Split (vValue, "~")

    vEcom_Quantity        = aValue(7)                       '...get the no of licenses from the end of the string, should be the same number if more than one program

    svProd_No             = svProd_No + 1
    svProd_Max            = svProd_Max + 1
    ReDim Preserve saProd (6, svProd_No)

    '...generate two values for each product

    saProd(0, svProd_No) = 0                              '...percentage discount
    saProd(1, svProd_No) = aValue(0)                      '...program Id
    saProd(2, svProd_No) = 1                              '...1 licenses
    saProd(3, svProd_No) = aValue(3)                      '...US price each
    saProd(4, svProd_No) = aValue(4)                      '...CA price each
    saProd(5, svProd_No) = aValue(5)                      '...no days duration
    saProd(6, svProd_No) = "<b>" & "<!--{{-->License<!--}}-->" & ":</b>&nbsp;" & aValue(6)


    svProd_No             = svProd_No + 1
    svProd_Max            = svProd_Max + 1
    ReDim Preserve saProd (6, svProd_No)

    saProd(0, svProd_No) = 0                              '...percentage discount
    saProd(1, svProd_No) = aValue(0)                      '...program Id
    saProd(2, svProd_No) = aValue(7)                      '...get no licenses from 7th value (only used in qOrderCatalogue)
    saProd(3, svProd_No) = aValue(1)                      '...US price each
    saProd(4, svProd_No) = aValue(2)                      '...CA price each
    saProd(5, svProd_No) = aValue(5)                      '...no days duration
    saProd(6, svProd_No) = "<b>" & aValue(7) & "&nbsp;" & "<!--{{-->Seats<!--}}-->" & ":</b>&nbsp;" & aValue(6)


    '...save basket values
    Session("ProdNo")    = svProd_No
    Session("ProdMax")   = svProd_Max
    Session("Prod")      = saProd       
  
  Next 
  
  Response.Redirect "Ecom2Customer.asp?vMode=More"

%>

