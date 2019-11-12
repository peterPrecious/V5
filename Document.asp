<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Document.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<%
  '...similiar to Certificate.asp
  Dim vFileName, vCustId, vProgId, vModsId, vLang, vMemo
  
' //vubiz.com/v5/document.asp?vCustId=VUBZ2277&vFileName=harassment.pdf&vMembNo=1034754&vProgId=P2362EN&vModsId=1630EN&vLang=EN

  vProgId     = ""
  vModsId     = ""
  vMemo       = ""

  vFileName   = Ucase(fDefault(Request("vFileName"), "Harrassment.pdf"))
  vCustId     = Request("vCustId")
  vLang       = Request("vLang")
    
  '...set these to ensure you can access the DB without signing in to get the Account ID (might be different than the last 4 digits of the Cust ID) 
  Session("CustId")     = vCustId
  Session("MembId")     = "WebService"
  Session("CustAcctId") = Right(vCustId, 4)
  Session("HostDb")     = "V5_Vubz"
  sGetCust vCustId

  Response.Redirect fDocumentUrl(vFileName, vModsId, vLang, Left(vCustId, 4), Right(vCust_Id, 4), vProgId, "")
%>