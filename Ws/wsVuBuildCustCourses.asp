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
       Session("HostDb") = "V5_Vubz"
       '...ensure customer is valid
       sGetCust (vCustId)
       If vCust_Eof Then 
         vErr = "vuBuild cannot continue due to an Account Setup Error.  Please contact your facilitator."
       Else
         vResponse = vResponse & vResponseXML
       End If

       If vErr = "" then
         '...read Prog for module names
         sReadCustMods
         If vCust_Eof Then 
           vErr = "vuBuild cannot continue due to a Module Setup Error.  Please contact your facilitator."
         Else
           vResponse = vResponse & vResponseXML
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
    vSql = "SELECT Cust_Id, Cust_Active, Cust_AcctId, Cust_Programs, Cust_Auth FROM Cust WHERE Cust_Id= '" & vCustId & "'"
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
  End Sub  

  Sub sReadCustMods
       sOpenDbBase
       vSql = "Select Prog_Mods FROM Prog WHERE Prog_Id = '" & vCust_Prgms & "'"
       Set oRs = oDbBase.Execute(vSQL)
       If Not oRs.Eof Then 
       vResponseXML = "   <Customer>" _
                    & "      <vCust_Mods>"        & oRs("Prog_Mods")   & "</vCust_Mods>" _
                    & "   </Customer>"
         vCust_Eof = False
       Else
         vCust_Eof = True
       End If
       Set oRs = Nothing
       sCloseDbBase
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