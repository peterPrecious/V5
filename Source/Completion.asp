<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  Dim vNext : vNext = fDefault(Request("vNext"), "Completion_0.asp")

  '...if user has changed accounts this session ...
  If svCustAcctId <> Session("Completion_AcctId") Then

    '...clear out all the session variables 
    For Each i In Session.Contents
      If Instr(i, "Completion_") > 0 Then
        Session(i) = ""
      End If
    Next

    '...initialize 
    Session("Completion_AcctId")        = svCustAcctId
    Session("Completion_Debug")         = False
    Session("Completion_InitParms")     = "Y"         '...control if/when to create session variables of all parameters (Y:initialize, N:initialized, not set:ignore)
    Session("Completion_InitContent")   = "Y"         '...control if/when to create the table of all content (Y:initialize, N:initialized, not set:ignore - reports)

    Session("Completion_checkRoles")    = "off"
    Session("Completion_checkProgs")    = "off"

  End If

  Response.Redirect "Patience.asp?vNext=" & vNext
%>
