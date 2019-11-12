 <!--#include virtual = "V5/Inc/Setup.asp"-->
 <%vBypassSecurity = True %>
 <!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  Dim vResponse, vResponseNamePair, oXmlHttp, vTmpNamePair, vResponseXML, vErr
  '...vResponseNamePair setup in the WS includes
  Dim vAction, vCustID, vMembID, vPass 

  '...Common - extract mandatory parms___________________________________________________
   vAction = Request("vAction")
   vCustID = Request("vCustID")
   vMembID = Request("vMembID")
   vPass   = Request("vPass")
   vErr    = ""
  '______________________________________________________________________________________

  If vErr = "" Then 
     '... Setup XML response
     vResponse = vResponse & "<?xml version='1.0' ?>" _
                        & "  <VUBUILD>"
     '... Grab Company info
     If vAction   = "GetCustRecs" then
       Session("HostDb") = "V5_Vubz"
       '...ensure customer is valid
       sGetCust (vCustId)
       If vCust_Eof Then 
         vErr = "vuBuild cannot log you in due to an Account Setup Error.  Please contact your facilitator."
       Else
         vResponse = vResponse & vResponseXML
       End If

       If vErr = "" then
         '...read Prog for module names
         sReadCustMods
         If vCust_Eof Then 
           vErr = "vuBuild cannot log you in due to a Module Setup Error.  Please contact your facilitator."
         Else
           vResponse = vResponse & vResponseXML
         End If
       End If

       If vErr = "" then
         '... Grab Company Members Info
         If Not vCust_Eof then 
           fMemb_List(vCust_AcctId)
           If vMemb_Eof then
             vErr = "vuBuild cannot log you in due to an Author Setup Error.  Please contact your facilitator."
           Else
             vResponse = vResponse & vResponseXML
           End If 
         End If
       End If  
    End If
  End If  
  
  '... Close Off XML
  vResponse = vResponse & "    <Error>" _
                        & "      <ErrorType>" & vErr & "</ErrorType>" _                        
                        & "    </Error>" _
                        & "  </VUBUILD>" 
  
  
  '...Return
  Response.Write vResponse
  
  '... Sub and Functions .............................................................
  Dim vCust_Id, vCust_AcctId, vCust_Title, vCust_Lang, vCust_Active, vCust_Auth, vCust_Prgms
  Dim vCust_Eof, vCust_Mods

  Sub sGetCust (vCustId)

    vSql = "SELECT Cust_Id, Cust_AcctID, Cust_Title, Cust_Lang, Cust_Programs, Cust_Active, Cust_Auth FROM Cust WHERE Cust_Id= '" & vCustId & "'"
    sOpenDB
    Set oRs = oDB.Execute(vSql)

    If Not oRs.Eof Then 
      sReadCust
      vCust_Eof = False
      If vCust_Active = 0 then vCust_Eof = True
      If vCust_Auth   = 0 then vCust_Eof = True
    Else
      vCust_Eof = True
    End If
    Set oRs = Nothing
    sCloseDB    
  End Sub
  Sub sReadCust
    vCust_Active = oRs("Cust_Active")
    vCust_Auth   = oRs("Cust_Auth")
    vCust_AcctId = oRs("Cust_AcctId")
    vCust_Prgms  = fStripOutPrgms(oRs("Cust_Programs"))
    vResponseXML =  "    <Customer>" _
                  & "      <vCust_Id>"      & Server.HTMLEncode(oRs("Cust_Id"))                  & "</vCust_Id>"      _
                  & "      <vCust_AcctId>"  & Server.HTMLEncode(oRs("Cust_AcctId"))              & "</vCust_AcctId>"  _
                  & "      <vCust_Title>"   & Server.HTMLEncode(oRs("Cust_Title"))               & "</vCust_Title>"   _
                  & "      <vCust_Lang>"    & oRs("Cust_Lang")                                   & "</vCust_Lang>"    _
                  & "      <vCust_Prgms>"   & Left(oRs("Cust_Programs"), 7)                      & "</vCust_Prgms>"   _
                  & "      <vCust_Active>"  & oRs("Cust_Active")                                 & "</vCust_Active>"  _
                  & "      <vCust_Auth>"    & oRs("Cust_Auth")                                   & "</vCust_Auth>"    
  End Sub  
  Sub sReadCustMods
       sOpenDbBase
       vSql = "Select Prog_Mods, Prog_Title, Prog_Desc FROM Prog WHERE Prog_Id = '" & vCust_Prgms & "'"
       Set oRs = oDbBase.Execute(vSQL)
       If Not oRs.Eof Then 
       vResponseXML = "      <vCust_Mods>"        & oRs("Prog_Mods")                      & "</vCust_Mods>"       _
                    & "      <vCust_Prog_Title>"  & Server.HTMLEncode(oRs("Prog_Title"))  & "</vCust_Prog_Title>" _
                    & "      <vCust_Prog_Desc>"   & Server.HTMLEncode(oRs("Prog_Desc"))   & "</vCust_Prog_Desc>"  _
                    & "    </Customer>" 
         vCust_Eof = False
       Else
         vCust_Eof = True
       End If
       Set oRs = Nothing
       sCloseDbBase
  End Sub

  '... Memb  .............................................................
  '...Return Member List string of FirstName, LastName, No, Email
  Dim vMemb_AcctId, vMemb_Id, vMemb_No, vMemb_FirstName, vMemb_LastName, vMemb_Email 
  Dim vMemb_Eof, vMemb_Level, vMemb_Auth
  Function fMemb_List(svCustAcctId)
    fMemb_List = ""
    '...ignore inactive and only get authors
    vSql = "SELECT * FROM Memb WHERE Memb_AcctId = '" & svCustAcctId & "' AND Memb_Active = 1 AND Memb_Auth = 1 AND Memb_Id = '" & vMembID & "'"
   'sDebug "vSql", vSql
    sOpenDb
    Set oRs = oDb.Execute(vSql)
    If Not oRs.Eof Then
      vMemb_Eof = False
      sReadMemb
    Else  
      vMemb_Eof = True
    End If
    Set oRs = Nothing
    sCloseDb
  End Function

  '...get the current fields from the current record in the record set
  Sub sReadMemb
    vResponseXML =  "    <Member>" _
                  & "      <vMemb_AcctId>"    & oRs("Memb_AcctId")                       & "</vMemb_AcctId>"      _
                  & "      <vMemb_No>"        & oRs("Memb_No")                           & "</vMemb_No>"          _
                  & "      <vMemb_FirstName>" & Server.HTMLEncode(oRs("Memb_FirstName")) & "</vMemb_FirstName>"   _
                  & "      <vMemb_LastName>"  & Server.HTMLEncode(oRs("Memb_LastName"))  & "</vMemb_LastName>"    _
                  & "      <vMemb_Email>"     & Server.HTMLEncode(oRs("Memb_Email"))     & "</vMemb_Email>"       _
                  & "      <vMemb_Level>"     & oRs("Memb_Level")                        & "</vMemb_Level>"       _
                  & "      <vMemb_Auth>"      & oRs("Memb_Auth")                         & "</vMemb_Auth>"        _
                  & "    </Member>" 

  End Sub
    

  Function fStripOutPrgms(vPrgmStr)
    fStripOutPrgms = ""
    Dim q:q=1
    Do While q < Len(vPrgmStr)
      If Mid(vPrgmStr, q, 1) = "P" then
        fStripOutPrgms = fStripOutPrgms & Mid(vPrgmStr, q, 7) & " "
        q = q + 7
      End If      
      q = q + 1    
    Loop
    fStripOutPrgms = Left(fStripOutPrgms, Len(fStripOutPrgms) - 1) '... strip out trailing space
  End Function  













%>