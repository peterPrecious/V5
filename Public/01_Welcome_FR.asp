<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<% Response.Redirect "/Chaccess/Signin" %>

<%  
  Dim vCust, vFrame
  vCust = Request("vCust")
  vFrame = fDefault(Request("vFrame"), "11_WhatsNew.asp")
  If Len(vCust) > 0 Then vFrame = vFrame & "?vCust=" & vCust
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
<!--      <li><a onmouseover="javascript:window.status=' ';return true" onmousedown="javascript:window.status=' ';return true" onmouseout="javascript:window.status=' ';return true" target="_top" href="/V5/Default.asp?vLang=FR">Français</a></li>-->          
          <li><a onmouseover="javascript:window.status=' ';return true" onmousedown="javascript:window.status=' ';return true" onmouseout="javascript:window.status=' ';return true" target="_top" href="/V5/Default.asp?vLang=ES">Espa�ol</a></li>
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
          <li><h3><a href="10_Login_FR.asp">Inscription</a></h3></li>
        </ul>
      </div>
      <div id="navcontainer3">
        <ul id="navlist3">
          <li><a target="_top" href="00_Home_FR.asp">Page d&#39;accueil</a></li>
        </ul>
      </div>
      <div id="navcontainer3">
        <ul id="navlist3">
          <li><a target="_top" href="Cat_Default.asp?vLang=FR&vCustId=VUBZ2275">Catalogue</a></li>
        </ul>
      </div>
      <div id="navcontainer3">
        <ul id="navlist3">
          <li><a href="BrowserIssues_FR.htm">Probl�mes li�s aux navigateurs?</a></li>
        </ul>
      </div>      

      <a target="_blank" title="Programme de certificat de strat�gies de marketing et ventes de petite entreprise" href="http://vubiz.com/chaccess/Certificate2011FR/">Certificat  MVPE</a>


      <p class="c2"><a target="_blank" href="http://gpeCertificat.com/default.asp?vMemo=VUBZ_P&vSource=http://vubiz.com"><img border="0" src="../Images/SPC/SBMC/SBMC_SM_FR.jpg" width="107" height="50"></a> <br /><br /><br>&nbsp;<a href="http://sspecertificat.com/default.asp?vSource=http://vubiz.com&vMemo=VUBZ_P"><img border="0" src="../Images/SPC/SBHS/SBHS_SM_FR.jpg" width="106" height="50"></a><br><br><a target="_blank" href="http://www.cchst.ca/education/pdf/TECatalogue.pdf">T�l�charger le catalogue d'�ducation et de formation du CCHST</a></td>
      <td valign="top" background="Home/Images/vubizFramed_r6_c3.jpg"><img name="vubizFramed_r6_c3" src="Home/Images/vubizFramed_r6_c3.jpg" width="9" height="526" border="0" alt=""></td>
      <td valign="top" bgcolor="#FFFFFF"><iframe src="<%=vFrame%>" name="iframe" width="566" height="520" scrolling="Auto" frameborder="0" id="iframe" target="self">[Your browser does not support frames or is currently configured not to display frames. However, you may visit <a href="new/00_WhatsNew.asp">the related document.</a>] 
        </iframe></td>
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