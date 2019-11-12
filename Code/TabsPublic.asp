<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<!-- This is used when LPs invoke the Catalogue, etc when not signed in -->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="<%=svDomain%>/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <link href="/V5/Inc/<%=Left(svCustId, 4)%>.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title></title>
</head>

<body>

  <table width="100%" height="54" border="0" cellpadding="0" cellspacing="0" background="../Images/Shell/TabsBg.gif">
    <tr>
      <td width="11" nowrap background="../Images/Shell/1x1TransparentSpacer.gif">
        <img src="../Images/Shell/1x1TransparentSpacer.gif" width="11" height="54"></td>
      <% If Len(svCustBanner) = 0 Then svCustBanner = "Vubz.jpg" %>
      <td width="45%" nowrap>
        <img border="0" src="/V5/Images/Logos/<%=svCustBanner%>"></td>
      <td valign="bottom" nowrap>&nbsp;</td>
      <td width="48%" valign="bottom" nowrap></td>
    </tr>
  </table>
  <table width="100%" border="0" cellspacing="0" cellpadding="0" id="table1">
    <tr>
      <td width="1%" valign="top">
        <img src="../Images/Shell/ActiveBar_TopRLeft.gif" width="23" height="22"></td>
      <td class="c1" align="center" background="../Images/Shell/ActiveBar_TopMiddle.gif" nowrap width="96%"></td>
      <td width="1%" valign="top">
        <img src="../Images/Shell/ActiveBar_TopRight.gif" width="23" height="22"></td>
      <td valign="bottom" nowrap width="1%" align="right">
        <img src="../Images/Shell/5x5.gif" width="16" height="22"></td>
    </tr>
  </table>

</body>

</html>




