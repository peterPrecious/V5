<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<% 

  '...Increment total number of minutes and flag user as being offline
  Dim vMin
  If Len(Session("CurrVisit")) > 0 Then
    vMin = (DateDiff("s", Session("CurrVisit"), Now())) \ 60
    If vMin < 2 Then vMin = 2
  	sOpenCmd
    With oCmd
      .CommandText = "spSignOff"
      .Parameters.Append .CreateParameter("@membNo",  adInteger, adParamInput, , Session("MembNo"))
      .Parameters.Append .CreateParameter("@minutes", adInteger, adParamInput, , vMin)
    End With
    oCmd.Execute()
    Set oCmd = Nothing
    sCloseDb
  End If

	Dim vSource, vLang, vLogo, vCust
	vSource	= Request.QueryString("vSource")

  vSource = Replace (vSource, "~1", "&")
  vSource = Replace (vSource, "~2", "=")
  vSource = Replace (vSource, "~3", "?")

  '... we need to process these values (set in Users_O.asp line 268) - in this case vGoto
  vSource = Replace (vSource, "~4", "~1")
  vSource = Replace (vSource, "~5", "~2")
  vSource = Replace (vSource, "~6", "~3")

	vCust 	= Request("vCust")
	vLang 	= Request("vLang")
	vLogo 	= Request("vLogo")

  '...if NOT trying to close the window then go to next page to consumate unless vSource is valid, NULL or CLOSE
  If vSource <> "CLOSE" Then

    '...signoff with no returning link
    If vSource = "NULL" Then
      Response.Redirect "SignOffOK.asp?vLang=" & Request.QueryString("vLang")  & "&vLogo=" & Request.QueryString("vLogo")   
 
    '...signoff and return to source
    ElseIf Len(Request.QueryString("vSource")) > 0 Then
			If Instr(vSource, "/CHACCESS") > 0 Then vSource = vSource & "?vCust=" & vCust & "&vLang=" & vLang
     '  Response.Redirect "/V5/Default.asp?vCust=VUBZ5678&vId=VUV5_ADM&vGoto=Default.asp~3vPage~2Users.asp"

     Response.Redirect vSource


    '...signoff and allow to reenter
    Else
      Response.Redirect "SignOffOK.asp?vCust=" & vCust & "&vLang=" & vLang  & "&vLogo=" & vLogo
    End If

  End If  
%>

<html>
<head>
  <meta charset="UTF-8">
  <title>Sign Off</title>
</head>
<body onload="window.close()"></body>
</html>


