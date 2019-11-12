<!--#include virtual = "V5\Inc\Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5\Inc\Initialize.asp"-->
<!--#include virtual = "V5\Inc\Db_Cust.asp"-->
<!--#include virtual = "V5\Inc\Db_Prog.asp"-->
<!--#include virtual = "V5\Inc\Db_Mods.asp"-->
<!--#include virtual = "V5\Inc\Certificate.asp"-->
<!--#include virtual = "V5\Code\ModuleStatusRoutines.asp"-->

<% 
  Dim vLastScore, vBestScore, vErrMsg

  sGetCust(svCustId)
  sGetMods(Request("vModsId"))
  vProg_Id = Left(Request("vProgId"), 5)

  vLastScore = fFormatDate(fLastScore (svMembNo, vMods_Id))
  vBestScore = fBestScore (svMembNo, vMods_Id)

  If vBestScore < 80 Then
    If svLang = "FR" Then
	    vErrMsg = "Pour générer un certificat,<br>vous devez obtenir au moins 80% dans cette évaluation."
    Else
      vErrMsg = "To generate a Certificate,<br>you must achieve at least 80% in this assessment."
    End If  
    Response.Redirect "/V5/Code/Error.asp?vClose=y&vReturn=close&vErr=" & Server.UrlEncode(vErrMsg)
  End If

  Response.Redirect fCertificateUrl("", "", vBestScore, vLastScore, vMods_Id, Server.URLEncode(vMods_Title), "", "", "", vProg_Id, "", "") 
%>

