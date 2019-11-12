<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Certificate.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<%
  '...this comes from vuAssess live or V8 - the other values come from the session variables
  Dim vProgId, vModsId, vScore, vEmail, vDate

  vProgId = fDefault(Ucase(Request("vProgId")), "P0000XX")  
  vModsId = Ucase(fIf(Len(Request("vModsId")) > 0, Request("vModsId"), Request("vModId")))
  vScore  = Request("vScore")
  vEmail  = Request("vEmail")
  vDate   = Request("vDate")

  '   on July 14 2014 we added the ability to generate certs for 3rd party users like CAMP who are not signed in
  Dim vCustId, vAcctId, vMembNo, vFirstName, vLastName, vMemo

  vCustId = Request("vCustId")
  vMembNo = Request("vMembNo")
  vMemo   = Request("vMemo")

  If (Session("Secure") = Empty And Len(vCustId)=8 And fPureInt(vMembNo) > 0) Then

    sOpenDb

    vSql = "SELECT Cust_AcctId FROM Cust WHERE Cust_Id = '" & vCustId & "'"
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      vAcctId = oRs("Cust_AcctId")
    End If
    Set oRs = Nothing      

    vSql = "SELECT Memb_FirstName, Memb_LastName FROM Memb WHERE Memb_No = " & vMembNo
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then 
      vFirstName  = oRs("Memb_FirstName")
      vLastName   = oRs("Memb_LastName")
    End If
    Set oRs = Nothing      

    sCloseDb
  End If



  '...log certificate (added Mar 22, 2016 to identify certs where prog didn't complete in RTE)
  Dim vIP
  vIP = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
  If vIP = "" Then vIP = Request.ServerVariables("REMOTE_ADDR")

  sOpenCmdApp
  With oCmdApp
    .CommandText = "sp5certLogs"
    .Parameters.Append .CreateParameter("@progId",  	      adVarChar, adParamInput,        7, vProgId)
    .Parameters.Append .CreateParameter("@modsId",  	      adVarChar, adParamInput,        8, vModsId)
    .Parameters.Append .CreateParameter("@score",  	        adInteger, adParamInput,         , fDefault(vScore, 0))
    .Parameters.Append .CreateParameter("@email",  	        adVarChar, adParamInput,      128, vEmail)
    .Parameters.Append .CreateParameter("@date",  	        adVarChar, adParamInput,       32, vDate)
    .Parameters.Append .CreateParameter("@custId",  	      adVarChar, adParamInput,        8, vCustId)
    .Parameters.Append .CreateParameter("@membNo",  	      adInteger, adParamInput,         , vMembNo)
    .Parameters.Append .CreateParameter("@memo",  	        adVarChar, adParamInput,      512, vMemo)
    .Parameters.Append .CreateParameter("@ip",  	          adVarChar, adParamInput,       32, vIP)
  End With
  oCmdApp.Execute()
	Set oCmdApp = Nothing
	sCloseDbApp

' Response.Redirect fCertificateUrl("", "", vScore, vDate, vModsId, fModsTitle(vModsId), "", "", "", vProgId, "", "", vEmail)
  Response.Redirect fCertificateUrl(vFirstName, vLastName, vScore, vDate, vModsId, fModsTitle(vModsId), "", vCustId, vAcctId, vProgId, "", vMemo, vEmail)
%>