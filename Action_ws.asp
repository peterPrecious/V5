<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
	Dim bOk, vAction, vStatus, vCustId, vLearnerId, vManagerId, vManagerNo
	
  vAction  		= Ucase(Request.Form("vAction"))
  vCustId  		= Ucase(Request.Form("vCust"))
  vLearnerId 	= Ucase(Request.Form("vLearnerId"))
  vManagerId 	= Ucase(Request.Form("vManagerId"))
  vManagerNo	= 0

  vStatus  = "Invalid Action"

  '...determine if the user ID and Password are valid when signing in
  If Instr("ACTIVATE INACTIVATE", vAction) > 0 Then

		vStatus = "Ok"

		'...ensure customer is valid
		If vStatus = "Ok" Then
			sGetCust vCustId
			If vCust_Eof Then vStatus = "Inv Customer"
		End If
		
		'...ensure manager is valid
		If vStatus = "Ok" Then
			sGetMembById vCust_AcctId, vManagerId
			If vMemb_Eof Then 
				vStatus = "Inv Manager"			
			ElseIf vMemb_Level < 4 Then 
				vStatus = "Inv Manager Level"
			Else
				vManagerNo = vMemb_No
			End If
		End If			

		'...ensure learner is valid
		If vStatus = "Ok" Then
			sGetMembById vCust_AcctId, vLearnerId
			If vMemb_Eof Then 
				vStatus = "Inv Learner"			
				If vMemb_Level <> 2 Then 
					vStatus = "Inv Learner Level"
				End If
			End If
		End If				


		'...allows manager to activate or inactive a learner whose learner id is passed in vSource
		spMembActiveById vCust_AcctId, vLearnerId, fIf(vAction = "INACTIVATE", 0, 1), vManagerNo

  End If  
  
  Response.Write vStatus
  
%>
