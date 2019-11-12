<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <link href="../V5/Inc/<%=Left(svCustId, 4)%>.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="Inc/RightClick.js"></script><% End If %>
  <title>:: Browser Test</title>
  <script>
    var t1, t2, t3, t4, t5, t6;
    var s1 = true, s2 = true, s3 = true, s4 = true, s5 = true, s6 = true;
    var jsver = 0.0;
  </script>
  <!-- Check Javascript version -->
  <script type="text/javascript">
    jsver = 1.0;
  </script>
  <script language="Javascript1.1">
    jsver = 1.1;
  </script>
  <script language="Javascript1.2">
    jsver = 1.2;
  </script>
  <script language="Javascript1.3">
    jsver = 1.3;
  </script>
  <script language="Javascript1.4">
    jsver = 1.4;
  </script>
  <script language="Javascript1.5">
    jsver = 1.5;
  </script>
  <script language="Javascript1.6">
    jsver = 1.6;
  </script>
  <script>

    function createCookie(name,value,days) {
      if (days) {
        var date = new Date();
        date.setTime(date.getTime()+(days*24*60*60*1000));
        var expires = '; expires='+date.toGMTString();
      }
      else var expires = '';
      document.cookie = name+'='+value+expires+'; path=/';
    }  
    function readCookie(name) {
      var nameEQ = name + '=';
      var ca = document.cookie.split(';');
      for(var i=0;i < ca.length;i++) {
        var c = ca[i];
        while (c.charAt(0)==' ') c = c.substring(1,c.length);
          if (c.indexOf(nameEQ) == 0) return c.substring(nameEQ.length,c.length);
      }
      return null;
    }  
    function eraseCookie(name) {
      createCookie(name,'',-1);
    }


    // check javascript
    t1 = "Javascript: " + jsver;
    s1 = true;

    // set negative in case not reviewed
    t3 = "Browser not analyzed.";
    t4 = "Flash Player not analyzed.";
    t5 = "Screen Size not analyzed.";
    t6 = "Cookies not analyzed.";


    // check platform
    if (navigator.userAgent.indexOf('Win') == -1) {
      s2 = false;
      t2 = "/*--{[--*/You must use Windows/*--]}--*/"
    }
    else {
      t2 = "Platform: Windows";

      // check browser
      var agt         = navigator.userAgent.toLowerCase(); 
      var is_ie       = (agt.indexOf("msie") != -1); 

      if (!is_ie) {
        s3 = false;
        t3 = "/*--{[--*/You must use Internet Explorer 6 or later/*--]}--*/";
      }
      else {
        var is_ie6    = (is_ie && (agt.indexOf("msie 6")!=-1)); 
        var is_ie7    = (is_ie && (agt.indexOf("msie 7")!=-1)); 
        var is_ie8    = (is_ie && (agt.indexOf("msie 8")!=-1)); 
        if (!is_ie6 && !is_ie7 && !is_ie8) {
          s3 = false;
          t3 = "/*--{[--*/You must use Internet Explorer 6 or later/*--]}--*/";
        }
        else {
          if (is_ie6) {
            t3 = "Browser: IE 6";
          }
          else if (is_ie7) {
            t3 = "Browser: IE 7";
          }
          else if (is_ie8) {
            t3 = "Browser: IE 8";
          }
        
             
          // Comprehensive Flash detection
          theVBScript =  '<form name="vbform"><input type="hidden" name="flashdetect"><input type="hidden" name="rpdetect">\</form>\n';
          theVBScript += '<SCR' + 'IPT LANGUAGE="VBScript">\n';
          theVBScript += 'on error resume next\n';
          theVBScript += 'RealPyr = "False"\n';
          theVBScript += 'Set oTest = CreateObject("ShockwaveFlash.ShockwaveFlash.5")\n';
          theVBScript += 'If Err = 0 Then FlashInstalled = 5\n';
          theVBScript += 'Err.Clear\n';
          theVBScript += 'Set oTest = CreateObject("ShockwaveFlash.ShockwaveFlash.6")\n';
          theVBScript += 'If Err = 0 Then FlashInstalled = 6\n';
          theVBScript += 'Err.Clear\n';
          theVBScript += 'Set oTest = CreateObject("ShockwaveFlash.ShockwaveFlash.7")\n';
          theVBScript += 'If Err = 0 Then FlashInstalled = 7\n';
          theVBScript += 'Err.Clear\n';
          theVBScript += 'Set oTest = CreateObject("ShockwaveFlash.ShockwaveFlash.8")\n';
          theVBScript += 'If Err = 0 Then FlashInstalled = 8\n';
          theVBScript += 'Err.Clear\n';
          theVBScript += 'Set oTest = CreateObject("ShockwaveFlash.ShockwaveFlash.9")\n';
          theVBScript += 'If Err = 0 Then FlashInstalled = 9\n';
          theVBScript += 'select case FlashInstalled\n';
          theVBScript += '  case 2,3,4,5,6,7,8,9\n';
          theVBScript += '    document.vbform.flashdetect.value = FlashInstalled\n';
          theVBScript += '  case else\n';
          theVBScript += '    FlashInstalled = 0\n';
          theVBScript += '    document.vbform.flashdetect.value = FlashInstalled\n';
          theVBScript += 'end select\n';
          theVBScript += 'document.vbform.rpdetect.value = RealPyr\n';
          theVBScript += '<\/SCRIPT>\n';
        	document.write(theVBScript);
        	flashversion = document.vbform.flashdetect.value;
          if (flashversion==0) {
            s4 = false;
            t4 = "/*--{[--*/You must install the Flash player/*--]}--*/";
        	}
        	else {
          	t4 = "Flash player: " + flashversion;
          }
      
          // check screen size
        	if (window.screen.width < 1024) {
            s5 = false;
            t5 = "/*--{[--*/You must use a 1024 screen size or larger/*--]}--*/";
        	}
        	else {  
          	t5 = 'Screen size: ' + window.screen.width + ' x ' + window.screen.height + ' pixels';
          }

       
          // check if cookies are enabled
          var vCookies = true
          createCookie('VuAssess','written',10)
          if (readCookie('VuAssess')==null) vCookies = false
          eraseCookie('VuAssess')
          if (vCookies) {
          	t6 = 'Cookies: enabled';
          }
          else {
            s6 = false;
            t6 = "/*--{[--*/You must enable Cookies in your browser/*--]}--*/";
          } 
        }  
      }
    }

//  alert(t1 + '\n' + t2 + '\n' + t3 + '\n' + t4 + '\n' + t5 + '\n' + t6);
//  alert(s1 + '\n' + s2 + '\n' + s3 + '\n' + s4 + '\n' + s5 + '\n' + s6);

    if (s1 && s2 && s3 && s4 && s5 && s6) {
      location.href = "FTTC_Login.asp"
    }  
  </script>
</head>

<body bgcolor="#003467">

  <table height="100%" cellspacing="0" cellpadding="0" width="100%" border="0">
    <tr>
      <td valign="center" align="middle" width="100%">
      <table cellspacing="0" cellpadding="0" border="0">
        <tr>
          <td height="73" bgcolor="#FFFFFF" nowrap align="center" width="388" class="c2">
          <table border="0" id="table1" cellspacing="0" cellpadding="0">
            <tr>
              <td align="center" class="c5">
              <!--[[-->Please Note...<!--]]--><br>
              <script>
                if (!s1) {document.write(t1)};
                if (!s2) {document.write(t2)};
                if (!s3) {document.write(t3)};
                if (!s4) {document.write(t4)};
                if (!s5) {document.write(t5)};
                if (!s6) {document.write(t6)};
              </script>
              <h2><b>
              <!--[[-->This Service requires:<!--]]--></b></h2>
              </td>
            </tr>
            <tr>
              <td>
              <ul class="c2">
                <li>
                <!--[[-->Windows XP or later<!--]]--></li>
                <li>
                <!--[[-->Internet Explorer 6 or later - English or French<!--]]--> (<a target="_blank" href="//microsoft.com/downloads/"><u><!--[[-->Download<!--]]--></u></a>)</li>
                <li>
                <!--[[-->Firefox 3 or later - English or French<!--]]--> (<a target="_blank" href="//firefox.com/"><u><!--[[-->Download<!--]]--></u></a>)</li>
                <li>
                <!--[[-->Flash player 9 or later<!--]]--> (<a target="_blank" href="///downloads/"><u><!--[[-->Download<!--]]--></u></a>)</li>
                <li>
                <!--[[-->1024 screen size minimum<!--]]--></li>
                <li>
                <!--[[-->Cookies and Javascript enabled in browser<!--]]--></li>
              </ul>
              </td>
            </tr>
            <tr>
              <td align="center" class="black">
              <% If svLang = "EN" Then %> 
                <a class="black" href="BrowserTest2.asp?vLang=FR">Fran�ais</a> 
              <% Else %> 
                <a href="BrowserTest2.asp?vLang=EN">English</a> 
              <% End If %> 
              </td>
            </tr>
          </table>
          </td>
        </tr>
      </table>
      </td>
    </tr>
  </table>

</body>

</html>
