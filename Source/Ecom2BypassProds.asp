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
  '   the last value is how they must arrive
  '   use this example to build the url which looks like this:
  '   /V5/default.asp?vCust=1234EN&vAction=qOrderProds&vMemo=P1014EN%7E79%7E125%7E79%7E125%7E90%7EFeelings%3A+Quality+Service%2E%2E%2EFirst+Time%2C+Every+Time+for+Professionals%7CP1017EN%7E79%7E125%7E79%7E125%7E90%7EFeelings%3A+Customer+Service+for+those+in+Service+Retail

  ' vMemo = "00000000_C1234567~79~125~79~125~90~Poker+Time+V1~2|" _
  '       & "00000000_C1234568~79~125~79~125~90~Poker+Time+V2~1"
  ' vMemo = Server.UrlEncode(vMemo)

  Dim aPrograms, aValue
  
  Session("Ecom_Media") = "Prods"
  
  sGetQueryString  '...this gets the vMemo field (with all the other fields)
  aPrograms = Split(vMemo, "|")

  '...put values in prod array
  svProdNo  = 0
  svProdMax = 0

  '...put all the selected items into an array
  For i = 0 to Ubound(aPrograms)
  
    vValue                = Replace(aPrograms(i), "'", " ")
    aValue                = Split (vValue, "~")

    svProdNo             = svProdNo + 1
    svProdMax            = svProdMax + 1

    ReDim Preserve saProd (6, svProdNo)

    saProd(0, svProdNo) = 0                              '...percentage discount
    saProd(1, svProdNo) = aValue(0)                      '...product Id
    saProd(2, svProdNo) = aValue(7)                      '...get quantity from 7th value (only used in qOrderGroup and qOrderProds)
    saProd(3, svProdNo) = aValue(1)                      '...US price each
    saProd(4, svProdNo) = aValue(2)                      '...CA price each
    saProd(5, svProdNo) = 0                              '...no days duration
    saProd(6, svProdNo) = aValue(7) & " of <b>" & aValue(6) & "</b>"  '...title

    '...save basket values
    Session("ProdNo")    = svProdNo
    Session("ProdMax")   = svProdMax
    Session("Prod")      = saProd       
  
  Next 
  
  Response.Redirect "Ecom2Customer.asp"
%>

