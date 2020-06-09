<%@  codepage="65001" %>

<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  Dim param, selected, vSelected, vCnt, vHtml

  param = Request("param")
  selected = Request("selected")

  vHtml = vbCrLf 
  vCnt = 0

  '... valid param?
  If Instr("PI|PT|MI|MT", param) = 0 Then Response.Write ("err")

  '...get the sorted IDs or Titles for this account
  sOpenCmdBase
  With oCmdBase
    .CommandText = "spHistory_ws"
    .Parameters.Append .CreateParameter("@Acct", adChar, adParamInput, 4, svCustAcctId)
    .Parameters.Append .CreateParameter("@Type", adChar, adParamInput, 2, param)
  End With
  Set oRsbase = oCmdBase.Execute()
  If Not oRsbase.Eof Then 

    '...create the dropdown for either Modules or Programs
    If Left(param, 1) = "M" Then
      Do While Not oRsBase.Eof 
        vSelected = fIf(Instr(selected, oRsBase("Mods_Id")) > 0, " selected ", "")
        vHtml = vHtml & "<option value='" & oRsBase("Mods_Id") & "'" & vSelected & ">" & oRsBase("Mods_Id") & " : " & oRsBase("Mods_Title") & "</option>" & vbCrLf
        vCnt  = vCnt + 1
        oRsBase.MoveNext
      Loop
      vCnt = fIf(vCnt > 50, 12, fIf(vCnt > 8, 8, vCnt))
      vHtml = "Leave unselected for ALL. Use Ctrl+Enter for multiple selections...<br><br><select id='modsList' style='width:500px' size='" & vCnt & "' name='vMods' multiple>" & vHtml & "</select>"
    Else
      Do While Not oRsBase.Eof 
        vSelected = fIf(Instr(selected, oRsBase("Prog_Id")) > 0, " selected ", "")
        vHtml = vHtml & "<option value='" & oRsBase("Prog_Id") & "'" & vSelected & ">" & oRsBase("Prog_Id") & " : " & oRsBase("Prog_Title1") & "</option>" & vbCrLf
        vCnt  = vCnt + 1
        oRsBase.MoveNext
      Loop
      vCnt = fIf(vCnt > 50, 12, fIf(vCnt > 8, 8, vCnt))
      vHtml = "Leave unselected for ALL. Use Ctrl+Enter for multiple selections...<br><br><select id='progList' style='width:500px' size='" & vCnt & "' name='vProg' multiple>" & vHtml & "</select>"
    End If

  End If

  Set oRsBase   = Nothing
  Set oCmdBase  = Nothing
  sCloseDbBase


  Response.Write (vHtml)

%>