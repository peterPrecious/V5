<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim vMembId
  vMembId = Ucase(Request.Form("vMembId"))

  '...set these to ensure you can access the DB without signing in 
  Session("CustId")     = "SCRM2802"
  Session("CustAcctId") = Right("SCRM2802", 4)
  Session("HostDb")     = "V5_Vubz"
  Session("MembId")     = vMembId

  vMemb_AcctId  = Session("CustAcctId")
  vMemb_Id      = Session("MembId")

  '...first see if VU_SCORMTRACKER is on member file, if not add it 
  sGetMembById Session("CustAcctId"), vMemb_Id
  If vMemb_Eof Then sAddMemb

  '...increase number of visits
  vSql = "UPDATE Memb SET Memb_NoVisits = Memb_NoVisits + 1 " _
       & ", Memb_LastVisit = '" & fFormatSqlDate(Now) & "'"  _
       & " WHERE Memb_AcctID = '" & Session("CustAcctId") & "' AND Memb_Id = '" & vMemb_Id & "'"
  sOpenDb
  oDb.Execute(vSql)
  sCloseDb

  '... Following line for writing back to browser, debugging. 
  'Response.Write "ok"
  'Session.Abandon 

%>
