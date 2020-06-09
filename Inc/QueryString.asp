<%
  '...these routines keep the basic parms needed to start the service
  '...they are held in a single session variable where you can get and put the values
  '...once the app is rolling (ie member signed in) this feature is not needed 
  '   as the values will be localized into the appropriate session variables

  '...typical value = "vCust=&vId=&vPwd=&vLang=&vFirstName=&vLastName=&vEmail=&vSource=//salesgorilla.com&vQModId=0004EN&zoneoffset=4&platform=win&browser=ie&bver=4&flashver=5"

  '...application variables
  Dim vHostDb, vCust, vId, vPwd, vLang, vFirstName, vLastName, vEmail, vMemo, vCriteria, vTraining, vErrUrl, vSource, vQModId, vQTestId, vGoTo, v8, vAction, vTskH_Id

  '...client variables from browser sniffer
  Dim vZoneOffset, vPlatform, vBrowser, vBVer, vFlashVer

  Sub sGetQueryString

    If Len(Session("QueryString")) = 0 Then 
      If vBypassSecurity Then 
        Exit Sub
      Else
        Response.Redirect "/V5/Code/Timeout.asp?vPage=" & Request.ServerVariables("Path_Info")
      End If
    End If

    Dim aVariable1, aVariable2
    aVariable1 = Split(Session("QueryString"), "&") 
    For i = 0 To Ubound(aVariable1)
      aVariable2 = Split(aVariable1(i), "=")
      Select Case aVariable2(0)
        Case "vHostDb"     : vHostDb     = aVariable2(1)
        Case "vCust"       : vCust       = aVariable2(1)
        Case "vId"         : vId         = aVariable2(1)
        Case "vPwd"        : vPwd        = aVariable2(1)
        Case "vLang"       : vLang       = aVariable2(1)
        Case "vFirstName"  : vFirstName  = aVariable2(1)
        Case "vLastName"   : vLastName   = aVariable2(1)
        Case "vEmail"      : vEmail      = aVariable2(1)
        Case "vMemo"       : vMemo       = aVariable2(1)
        Case "vTraining"   : vTraining   = aVariable2(1)
        Case "vCriteria"   : vCriteria   = aVariable2(1)
        Case "vErrUrl"     : vErrUrl     = aVariable2(1)
        Case "vSource"     : vSource     = aVariable2(1)
        Case "vQModId"     : vQModId     = aVariable2(1)
        Case "vQTestId"    : vQTestId    = aVariable2(1)
        Case "vGoTo"       : vGoTo       = aVariable2(1)
        Case "v8"          : v8          = aVariable2(1)
        Case "vAction"     : vAction     = aVariable2(1)
        Case "vTskH_Id"    : vTskH_Id    = aVariable2(1)

        Case "zoneoffset"  : vZoneOffset = aVariable2(1)
        Case "platform"    : vPlatform   = aVariable2(1)
        Case "browser"     : vBrowser    = aVariable2(1)
        Case "bver"        : vBVer       = aVariable2(1)
        Case "flashver"    : vFlashVer   = aVariable2(1)
      End Select
    Next 
    '...hold browser type and flash status in separate session variables
    Session("Browser") = vBrowser 
    '...if flashver in URL then make either "true" or "false" (for javascript)
    Session("Flash")    = "false"
    If VarType(vFlashVer) > 1 Then
      If IsNumeric(vFlashVer) Then
        If vFlashVer> 0 Then 
          Session("Flash") = "true"
        End If
      End If
    End If
  End Sub  


  '...this puts the individual values into the single session variable 
  Sub sPutQueryString
    Session("QueryString") =  "vHostDb=" & vHostDb & "&vCust=" & vCust & "&vId=" & vId & "&vPwd=" & vPwd & "&vLang=" & vLang & "&vFirstName=" & vFirstName & "&vLastName=" & vLastName & "&vEmail=" & vEmail & "&vMemo=" & vMemo & "&vTraining=" & vTraining & "&vCriteria=" & vCriteria & "&vSource=" & vSource & "&vQModId=" & vQModId & "&vQTestId=" & vQTestId & "&vGoTo=" & vGoTo & "&v8=" & v8 & "&zoneoffset=" & vZoneoffset & "&platform=" & vPlatform & "&browser=" & vBrowser & "&bver=" & vBVer & "&flashver=" & vFlashVer & "&vAction=" & vAction  & "&vTskH_Id=" & vTskH_Id
  End Sub  

  '...this gets a real querystring when session is abandoned and puts into the single session variable 
  Sub sReadQueryString
    vHostDb     = Request.QueryString("vHostDb")
    vCust       = Request.QueryString("vCust")
    vId         = Request.QueryString("vId")
    vPwd        = Request.QueryString("vPwd")
    vFirstName  = Request.QueryString("vFirstName")
    vLastName   = Request.QueryString("vLastName")
    vEmail      = Request.QueryString("vEmail")
    vMemo       = Request.QueryString("vMemo")
    vTraining   = Request.QueryString("vTraining")
    vCriteria   = Request.QueryString("vCriteria")
    vAction     = Request.QueryString("vAction")
    vLang       = Request.QueryString("vLang")
    vErrUrl     = Request.QueryString("vErrUrl")
    vSource     = Request.QueryString("vSource")
    vQModId     = Request.QueryString("vQModId")
    vQTestId    = Request.QueryString("vQTestId")
    vGoto       = Request.QueryString("vGoTo")
    v8          = Request.QueryString("v8")
    vAction     = Request.QueryString("vAction")
    vTskH_Id    = Request.QueryString("vTskH_Id")

    vZoneOffset = Request.QueryString("zoneoffset")
    vPlatform   = Request.QueryString("platform")
    vBrowser    = Request.QueryString("browser")
    vBVer       = Request.QueryString("bver")
    vFlashVer   = Request.QueryString("flashver")
    
    sPutQueryString
  End Sub  

  Sub sReadQueryStringForm
    Session("QueryString") =  Request.Form("vQueryString")
  End Sub  

  Sub sDebugQueryString
    Response.Write "<br><b><font color='ORANGE'>... " & Session("QueryString") & "</font></b>"
    Response.Flush
  End Sub

%>