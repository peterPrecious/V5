<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/QueryString.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Logs.asp"-->

<%
  Dim aTab, vTab, vTabActive, vImageActive, vMain, vModId, vOnLoadScript, vHomeUrl, MyWorldLaunch, aTabName, vGold

  sGetQueryString
  
  sGetCust svCustId  '...setup the tabs for specific customer

  sGetMemb svMembNo

  '...setup default banner logos, urls and email
  If Len(svCustReturnUrl) > 0 Then
    vHomeUrl = svCustReturnUrl
  Else
    vHomeUrl = "//" & svHost
  End If  
  If fNoValue(svCustBanner) Then 
    svCustBanner = "vubz.jpg"
  End If


  Redim aTab(2, 0)
  vTab = -1 '...this is used to build the tab set

  '...if postback pickup tab number to display as active, else set to -1 (all inactive)
  vTabActive = Request.QueryString("vTab")
  If Len(vTabActive) = 0 Then
    vTabActive = 0
  Else
    vTabActive = Cint(vTabActive)
  End If
  Session("TabActive") = True '...use to see if tabs used for shell purposes

  '...if postback create onload script for the frames
  vMain = Request.QueryString("vMain")
  '...decode url
  vMain = Replace(vMain, "~1", "&")
  vMain = Replace(vMain, "~2", "=")
  vMain = Replace(vMain, "~3", "?")

  If Not fNoValue(vMain) Then 
    vOnLoadScript = "onLoad=" & chr(34) 
    vOnLoadScript = vOnLoadScript & "parent.frames.main.location.href='" & vMain & "';" 
    vOnLoadScript = vOnLoadScript & chr(34)
  Else
    vOnLoadScript = ""
  End If
 
 
  Sub sTabName (vTab, vTabName)
    If Len(vTabName) > 0 Then
      aTabName = Split(vTabName, "|")
      If svLang = "EN" And Ubound(aTabName) >= 0 Then 
      	If Len(aTabName(0)) > 0 Then aTab(1, vTab) = Trim(Left(aTabName(0) & Space(20), 20))
      End If
      If svLang = "FR" And Ubound(aTabName) >= 1 Then
      	If Len(aTabName(1)) > 0 Then aTab(1, vTab) = Trim(Left(aTabName(1) & Space(20), 20))
      End If
      If svLang = "ES" And Ubound(aTabName) >= 2 Then
        If Len(aTabName(2)) > 0 Then aTab(1, vTab) = Trim(Left(aTabName(2) & Space(20), 20))
      End If
    End If
  End Sub


  If vCust_Tab1 Then 
    vTab = vTab + 1
    Redim Preserve aTab(2, vTab)
    aTab(0, vTab) = "Patience.asp?vNext=Default.asp"
    aTab(1, vTab) = fPhraH(000160)
    sTabName vTab, vCust_Tab1Name
    aTab(2, vTab) = "_top"  
  End If


  If vCust_Tab2 And svCustLevel > 2 Then 
    vTab = vTab + 1
    Redim Preserve aTab(2, vTab)
    MyWorldLaunch = "MyWorld.asp"
    '...If there's an intro page, then launch that first
    If Len(vCust_MyWorldLaunch) > 4 Then
      MyWorldLaunch = "/V5/Repository/" & svHostDb & "/" & svCustAcctId & "/Tools/" & vCust_MyWorldLaunch
    End If    
    aTab(0, vTab) = "TabsLive.asp?vMain=Patience.asp?vNext=" & MyWorldLaunch & "&vTab=" & vTab
    aTab(1, vTab) = fPhraH(000183)
    sTabName vTab, vCust_Tab2Name
    aTab(2, vTab) = "_self"  
  End If


  If vCust_Tab3 Then 
    vTab = vTab + 1
    Redim Preserve aTab(2, vTab)
    aTab(0, vTab) = "TabsLive.asp?vMain=Patience.asp?vNext=RTE_MyContent.asp&vTab=" & vTab
    aTab(1, vTab) = fPhraH(000182)
    sTabName vTab, vCust_Tab3Name
    aTab(2, vTab) = "_self"  
  End If


  If vCust_Tab4 And fTab4Ok Then 
    vTab = vTab + 1
    Redim Preserve aTab(2, vTab)

    '...Discussion Forum
    If vCust_Tab4Type = "CL" Then
        aTab(0, vTab) = "TabsLive.asp?vMain=Patience.asp?vNext=MyClasses.asp&vTab=" & vTab
        aTab(1, vTab) = fPhraH(001371)
        aTab(2, vTab) = "_self"  
    '...Discussion Forum
    ElseIf vCust_Tab4Type = "DF" Then
        aTab(0, vTab) = "TabsLive.asp?vMain=Patience.asp?vNext=Discussion.asp&vTab=" & vTab
        aTab(1, vTab) = fPhraH(001198)
        aTab(2, vTab) = "_self"  
    ElseIf vCust_Tab4Type = "SC" Then
        aTab(0, vTab) = "TabsLive.asp?vMain=Patience.asp?vNext=Scheduler.asp&vTab=" & vTab
        aTab(1, vTab) = fPhraH(001251)
        aTab(2, vTab) = "_self"  
    Else
      If svMembLevel < 3 Then 
        aTab(0, vTab) = "TabsLive.asp?vMain=Patience.asp?vNext=RC_Home.asp&vTab=" & vTab
        aTab(1, vTab) = fPhraH(000778)
        aTab(2, vTab) = "_self"  
      Else
        aTab(0, vTab) = "TabsLive.asp?vMain=Patience.asp?vNext=RC_Home.asp&vTab=" & vTab
        aTab(1, vTab) = fPhraH(000779)
        aTab(2, vTab) = "_self"  
      End If
    End If
    sTabName vTab, vCust_Tab4Name
  End If

  If vCust_Tab5 And ((svMembLevel = 2 And Len(Trim(fOkValue(vCust_ParentId))) = 0) Or (svMembLevel > 2) Or (vAction = "ECOMBYPASS")) Then 
    vTab = vTab + 1   
    Redim Preserve aTab(2, vTab)
    aTab(0, vTab) = "TabsLive.asp?vMain=Patience.asp?vNext=Ecom2Start.asp~3vMode~2More~1vContentOptions~2" & vContentOptions & "&vTab=" & vTab
    aTab(1, vTab) = fPhraH(000180)
    sTabName vTab, vCust_Tab5Name
    aTab(2, vTab) = "_self"  
  End If

  If (vCust_Tab6 And svMembLevel > 2) Or svMembLevel > 4 Then 
    vTab = vTab + 1
    Redim Preserve aTab(2, vTab)

    aTab(0, vTab) = "TabsLive.asp?vMain=Patience.asp?vNext=Menu.asp&vTab=" & vTab


    aTab(1, vTab) = fPhraH(000065)
    sTabName vTab, vCust_Tab6Name
    aTab(2, vTab) = "_self"  

    '...if program accepted "vTab=9" then set this tab to active (ie like Customer.asp)
    If vTabActive = 9 Then
      vTabActive = vTab
    End If
  End If

  If vCust_Tab7 Then 

    '...replaced Aug 24, 2017 as this did not handle SSL properly
    '   If Instr(svCustReturnUrl, "//www.vubiz.com/") > 0 Then svCustReturnUrl = Mid(svCustReturnUrl, 21)
    '   If Instr(svCustReturnUrl, "//vubiz.com/") > 0 Then svCustReturnUrl = Mid(svCustReturnUrl, 17)
       
    '...if local url then simplify (new version)
    i = Instr(svCustReturnUrl, "localhost/") : If i > 0 Then svCustReturnUrl = Mid(svCustReturnUrl, i + 9)
    i = Instr(svCustReturnUrl, "vubiz.com/") : If i > 0 Then svCustReturnUrl = Mid(svCustReturnUrl, i + 9)

    vTab = vTab + 1
    Redim Preserve aTab(2, vTab)
    aTab(0, vTab) = "SignOff.asp?vCust=" & svCustId & "&vLang=" & svLang & "&vLogo=" & svCustBanner & "&vSource=" & svCustReturnUrl
    aTab(1, vTab) = fPhraH(000240)
    sTabName vTab, vCust_Tab7Name
    aTab(2, vTab) = "_top" 
  End If

  Function fTab4Ok
    '...assume it's ok to display Tab4 
    fTab4Ok = True
  End Function 

%>

<html>

  <head>
    <title>TabsLive</title>
    <meta charset="UTF-8">
    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
    <style>
      html, body, div, table, tr, th, td, p   { font-family: "Segoe UI", Arial, Helvetica, sans-serif; font-size: 14px; }
      A                                       { text-decoration: none; }
      A:hover                                 { text-decoration: underline; }
    </style>
    <link href="/V5/Inc/Vubiz.css" rel="stylesheet">
  </head>

  <body <%=vonloadscript%> leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" link="#000080" vlink="#000080" alink="#000080" text="#000080">

    <table width="100%" height="54" border="0" cellpadding="0" cellspacing="0" background="../Images/Shell/TabsBg.gif">
      <tr>
        <td width="11" nowrap background="../Images/Shell/1x1TransparentSpacer.gif">
          <img src="../Images/Shell/1x1TransparentSpacer.gif" width="11" height="54"></td>
        <td width="45%" nowrap>
          <% If svCustURL <> "" And Instr(Lcase(svCustUrl), "vubiz") = 0 Then %>
          <a <%=fStatX%> href="//<%=svCustURL%>" target="_blank">
            <img border="0" src="../Images/Logos/<%=svCustBanner%>"></a>
          <% Else %>
          <img border="0" src="../images/Logos/<%=svCustBanner%>">
          <% End If %> 
        </td>

        <!-- create the tabs - vTab is the active tab which appears as light blue -->
        <%
        For i = 0 to Ubound(aTab, 2)
          '...if active use different coloured tab (same name but remove the "In" before "Active")
          If i = vTabActive Then
            vImageActive = "" 
          Else
            vImageActive = "In"       
          End If  
        %>
        <td valign="bottom">
          <table cellspacing="0" cellpadding="0" width="100%" border="0">
            <tr>
              <td valign="bottom" align="right" width="1%">
                <img src="../Images/Shell/TabLeft_<%=vImageActive%>Active.gif" border="0"></td>
              <td nowrap align="middle" width="98%" background="../Images/Shell/TabBg_<%=vImageActive%>Active.gif">
                <img height="1" src="../Images/Shell/1x1TransparentSpacer.gif" width="50" border="0"><br>

                <% If Application("Alert") = "y" And i = 0 And vCust_Tab1 Then %>

                <a <%=fStatX%> target="<%=aTab(2, i)%>" href="<%=aTab(0, i)%>"><%=aTab(1, i)%></a> !

                <% Else %>

                <a <%=fStatX%> target="<%=aTab(2, i)%>" href="<%=aTab(0, i)%>"><%=aTab(1, i)%></a>

                <% End If %>

              </td>
              <td align="left" width="1%">
                <img src="../Images/Shell/TabRight_<%=vImageActive%>Active.gif" border="0"></td>
            </tr>
          </table>
        </td>
        <%
        Next
        %>
        <td nowrap align="right" width="48%">
          <%  
            If Application("Alert") = "y" Then
            ' And svMembLevel = 5 
            Dim vAlert, vMsg
            If svLang = "FR" Then
              vMsg = "ALERTE! S'il vous plaît cliquer."
              vAlert = "Ce service sera interrompu pour fin d’amélioration et ne sera pas disponible le samedi 23 mai de 6h00 à 9h00 HNE. Nous nous excusons des inconvénients causés."
            ElseIf svLang = "ES" Then
              vMsg = "ALERTA! Por favor, haga clic en."
              vAlert = "Este servicio estará en mantenimiento de rutina y no estará disponible el sábado 23 de mayo 06 a.m.-09 a.m. EST. Nos disculpamos por cualquier inconveniente."
            Else
              vMsg = "ALERT! Please click."
              vAlert = "This service will be undergoing routine maintenance and will not be available on Saturday May 23rd from 6 am to 9 am EST. We apologize for any inconvenience."
            End If
          %>
          <div style="text-align:center; width:200px; margin:auto; background-color:yellow;"><a onclick="alert('<%=vAlert%>')" href="#"><%=vMsg%></a></div>
          <% Else %>        
        &nbsp;
      <% End If %>
        </td>
      </tr>
    </table>

    <table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="1%" valign="top"><img src="../Images/Shell/ActiveBar_TopRLeft.gif" width="23" height="22"></td>
        <td class="c1" width="96%" align="center" background="../Images/Shell/ActiveBar_TopMiddle.gif" nowrap></td>
        <td width="1%" valign="top"><img src="../Images/Shell/ActiveBar_TopRight.gif" width="23" height="22"></td>
        <td valign="bottom" nowrap width="1%" align="right"><img src="../Images/Shell/5x5.gif" width="16" height="22"></td>
      </tr>
    </table>

  </body>
</html>


