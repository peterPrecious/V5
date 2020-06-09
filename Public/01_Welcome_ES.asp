<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<% Response.Redirect "/Chaccess/Signin" %>

<%  
  Dim vFrame
  vFrame = fDefault(Request("vFrame"), "11_WhatsNew.asp")
%>

<html>

<head>
  <title>:: Vubiz</title>
  <meta charset="UTF-8">
  <link href="Home/Css/css.css" rel="stylesheet" type="text/css" />
  <base target="iframe">
</head>

<body bgcolor="#ffffff">

  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="771">
    <tr>
      <td><img src="Home/Images/spacer.gif" width="16" height="1" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="155" height="1" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="9" height="1" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="566" height="1" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="9" height="1" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="16" height="1" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="1" height="1" border="0" alt=""></td>
    </tr>
    <tr>
      <td rowspan="8"><img name="vubizFramed_r1_c1" src="Home/Images/vubizFramed_r1_c1.jpg" width="16" height="835" border="0" alt=""></td>
      <td colspan="4" valign="bottom" background="Home/Images/vubizFramed_r1_c2.jpg">
      <div id="navcontainer2" align="right">
        <ul id="navlist">
          <li><a onmouseover="javascript:window.status=' ';return true" onmousedown="javascript:window.status=' ';return true" onmouseout="javascript:window.status=' ';return true" target="_top" href="/V5/Default.asp?vLang=EN">English</a></li>
          <li><a onmouseover="javascript:window.status=' ';return true" onmousedown="javascript:window.status=' ';return true" onmouseout="javascript:window.status=' ';return true" target="_top" href="/V5/Default.asp?vLang=FR">Français</a></li>
<!--      <li><a onmouseover="javascript:window.status=' ';return true" onmousedown="javascript:window.status=' ';return true" onmouseout="javascript:window.status=' ';return true" target="_top" href="/V5/Default.asp?vLang=ES">EspaÃ±ol</a></li>-->
        </ul>
      </div>
      </td>
      <td rowspan="8"><img name="vubizFramed_r1_c6" src="Home/Images/vubizFramed_r1_c6.jpg" width="16" height="835" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="1" height="67" border="0" alt=""></td>
    </tr>
    <tr>
      <td colspan="4"><img name="vubizFramed_r2_c2" src="Home/Images/vubizFramed_r2_c2.jpg" width="739" height="10" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="1" height="10" border="0" alt=""></td>
    </tr>
    <tr>
      <td colspan="4"><img name="vubizFramed_r3_c2" src="Home/Images/vubizFramed_r3_c2.jpg" width="739" height="125" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="1" height="125" border="0" alt=""></td>
    </tr>
    <tr>
      <td colspan="4"><img name="vubizFramed_r4_c2" src="Home/Images/vubizFramed_r4_c2.jpg" width="739" height="9" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="1" height="9" border="0" alt=""></td>
    </tr>
    <tr>
      <td colspan="4"><img name="vubizFramed_r5_c2" src="Home/Images/vubizFramed_r5_c2.jpg" width="739" height="12" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="1" height="12" border="0" alt=""></td>
    </tr>
    <tr>
      <td align="center" valign="top" bgcolor="#FFFFFF">
      <div align="left"></div>

      <div id="navcontainer3">
        <ul id="navlist3">
          <li><h3><a href="10_Login_ES.asp">Conexión</a></h3></li>
        </ul>
      </div>

      <div id="navcontainer3">
        <ul id="navlist3">
          <li><a target="_top" href="00_Home_ES.asp">Base</a></li>
        </ul>
      </div>

      <div id="navcontainer3">
        <ul id="navlist3">
          <li><a target="_top" href="Cat_Default.asp?vLang=ES&vCustId=VUBZ2294">Catálogo</a></li>
        </ul>
      <div id="navcontainer4">
        <ul id="navlist4">
          <li><a  href="BrowserIssues_ES.htm">Problemas con tu navegador?</a></li>
        </ul>
      </div>
      </div>


      </td>
      <td valign="top" background="Home/Images/vubizFramed_r6_c3.jpg"><img name="vubizFramed_r6_c3" src="Home/Images/vubizFramed_r6_c3.jpg" width="9" height="526" border="0" alt=""></td>
      <td valign="top" bgcolor="#FFFFFF">
        <iframe src="<%=vFrame%>" name="iframe" width="566" height="520" scrolling="Auto" frameborder="0" id="iframe" target="self">
        [Your browser does not support frames or is currently configured not to display frames. However, you may visit <a href="new/00_WhatsNew.asp">the related document.</a>] 
        </iframe>
      </td>
      <td valign="top" background="Home/Images/vubizFramed_r6_c5.jpg"><img name="vubizFramed_r6_c5" src="Home/Images/vubizFramed_r6_c5.jpg" width="9" height="526" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="1" height="526" border="0" alt=""></td>
    </tr>
    <tr>
      <td colspan="4"><img name="vubizFramed_r7_c2" src="Home/Images/vubizFramed_r7_c2.jpg" width="739" height="70" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="1" height="70" border="0" alt=""></td>
    </tr>
    <tr>
      <td colspan="4"><img name="vubizFramed_r8_c2" src="Home/Images/vubizFramed_r8_c2.jpg" width="739" height="16" border="0" alt=""></td>
      <td><img src="Home/Images/spacer.gif" width="1" height="16" border="0" alt=""></td>
    </tr>
  </table>
  </center>

</body>

</html>
