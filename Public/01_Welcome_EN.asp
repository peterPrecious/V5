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
  <link href="Home/Css/css.css" rel="stylesheet" type="text/css">
  <link href="http://vubiz.com/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script>
    function jSampler(modid) {
      var url = "http://vubiz.com/V5/Default.asp?vLang=EN&vCust=DEMO1001&vId=VUDEM&scorm=0&vClose=y&vQModId="+modid;
      modwindow = window.open(url,'Module','toolbar=no,location=1,width=785,height=575,left=50,top=50,status=yes,scrollbars=yes,resizable=yes');
    }
  </script>  
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
<!--      <li><a onmouseover="javascript:window.status=' ';return true" onmousedown="javascript:window.status=' ';return true" onmouseout="javascript:window.status=' ';return true" target="_top" href="/V5/Default.asp?vLang=EN">English</a></li>-->
          <li><a onmouseover="javascript:window.status=' ';return true" onmousedown="javascript:window.status=' ';return true" onmouseout="javascript:window.status=' ';return true" target="_top" href="/V5/Default.asp?vLang=FR">Français</a></li>
          <li><a onmouseover="javascript:window.status=' ';return true" onmousedown="javascript:window.status=' ';return true" onmouseout="javascript:window.status=' ';return true" target="_top" href="/V5/Default.asp?vLang=ES">Español</a></li>
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
      <br><input onclick="iframe.location.href='10_login_EN.asp?vCust=<%=vCust%>'" type="button" value="Login" name="bLogin" id="bButton" class="button100"><div id="navcontainer">
        <ul id="navlist">
          <li><br><a target="_top" href="00_Home_EN.asp">Home</a></li>
          <li><a target="_top" title="Main Course Catalogue" href="/V5/Default.asp?vLang=EN&vCustId=VUBZ2274&vAction=ORDER">Catalogue</a></li>
          <li><a target="_blank" title="Small Business Marketing and Sales Certificate Program" href="http://vubiz.com/chaccess/Certificate2011/">&nbsp;- SBMS Certificate</a></li>
          <li><a target="_blank" title="Small Business Certificate" href="http://sbmcertified.com/default.asp?vMemo=VUBZ_P&vSource=http://vubiz.com">&nbsp;- SBM Certificate</a></li>
          <li><a target="_blank" title="Small Business Health and Safety Certificate" href="http://sbhscertificate.com/default.asp?vMemo=VUBZ_P&vSource=http://vubiz.com">&nbsp;- SBHS Certificate</a></li>


          <li><a target="_blank" title="Small Business Human Resources Certificate" href="http://vubiz.com/chaccess/SBHR-US/">&nbsp;- SBHR Certificate (US)</a></li>
          <li><a target="_blank" title="Small Business Human Resources Certificate" href="http://vubiz.com/chaccess/SBHR/">&nbsp;- SBHR Certificate (CA)</a></li>
          <li><a target="_blank" title="Human Resources Generalist Certificate [Worth 16 HRCI credits]" href="http://vubiz.com/v5/default.asp?vCust=ERGP2962&vAction=ORDER&vTraining=35426">&nbsp;- HR Generalist Certificate</a></li>


          <li><a target="_blank" href="#" onclick="jSampler('3860EN')"><font color="Red">Making Your Purchase</font></a></li>
          <li><a href="11_WhatsNew.asp">What&#39;s New?</a></li>
          <li><a href="12_AboutVubiz.asp">About Vubiz</a></li>
          <li><a href="13_ProductsAndServices.asp">Products and Services</a></li>
          <li><a href="14_HowWeWorkWithYou.asp">How We Work With You</a></li>
          <li><a href="15_CorporateGovernance.asp">Corporate Governance</a></li>
          <li><a href="16_SuccessStories.asp">Success Stories</a></li>
          <li><a href="17_Awards.asp">Awards</a></li>
          <li><a href="18_MediaReleases.asp">Media Releases <font color="#FF0000">&nbsp; <i>NEW</i></font></a></li>
          <li><a href="19_ContactUs_US.asp">Contact Us - USA</a></li>
          <li><a href="20_ContactUs_CA.asp">Contact Us - Canada</a></li>
          <li><a href="AccessIssue.asp">Forget Your Password?</a></li>
          <li><a href="21_FAQ.asp">Technical Support | FAQ</a></li>
          <li><a href="22_PrivacyPolicy.asp">Privacy Policy</a></li>
          <li><a target="_blank" href="http://www.getabstract.com/servlets/Turnkey?u=nextmove">getAbstract</a></li>
        </ul>
      </div>
      <p class="c2">
      
      <a target="_blank" href="Members.asp">Vubiz is a proud member of these associations</a><br><br>
      <font color="#3366CC" size="2" face="Arial, Helvetica, sans-serif"><strong>Explore our new Certificate Programs !</strong></font><br><br>
      <a target="_blank" href="http://sbmcertified.com/default.asp?vMemo=VUBZ_P&vSource=http://vubiz.com"><img border="0" src="../Images/SPC/SBMC/SBMC_SM_EN.jpg" width="106" height="50"></a><br><br>
      <a target="_blank" href="http://sbhscertificate.com/default.asp?vMemo=VUBZ_P&vSource=http://vubiz.com"><img border="0" src="../Images/SPC/SBHS/SBHS_SM_EN.jpg" width="106" height="50"></a><br><br>
         
      <img onclick="alert('Please select the version below.')" border="0" src="../Images/SPC/SBHR/SBHR.jpg" width="130" height="79"><br>
      <a target="_blank" class="c2" href="http://vubiz.com/chaccess/SBHR-US/">SBHR Certificate (US)</a><br>
      <a target="_blank" href="http://vubiz.com/chaccess/SBHR/">SBHR Certificate (CA)</a><br><br>

      <a target="_blank" href="http://www.ccohs.ca/education/pdf/TECatalogue.pdf"><img border="0" src="Images/CCOHS.jpg" width="122" height="29"></a><br><br>
      <font color="#3366CC"><a target="_blank" class="c2" href="http://www.ccohs.ca/education/pdf/TECatalogue.pdf">Download the CCOHS Training &amp; Education Catalogue</a></font><br>
      <a target="_blank" href="https://www.cga-pdnet.org/en-CA/Pages/default.aspx"><img border="0" src="Images/PDnet_1.jpg" width="100" height="56"></a><br><a target="_blank" href="http://www.vubiz.com/chaccess/padm_certificate"><img border="0" src="../Images/SPC/ICSA/ICSA_PAD_SM.jpg" width="104" height="50"></a>
      
      </p>

      
      </td>



      <td valign="top" background="Home/Images/vubizFramed_r6_c3.jpg"><img name="vubizFramed_r6_c3" src="Home/Images/vubizFramed_r6_c3.jpg" width="9" height="526" border="0" alt=""></td>
      <td valign="top" bgcolor="#FFFFFF">
        <iframe src="<%=vFrame%>" name="iframe" width="566" height="659" scrolling="Auto" frameborder="0" id="iframe" target="self">
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

  <script>document.getElementById("bButton").focus()</script>

</body>

</html>