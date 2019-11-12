<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  Dim vUrl
  If svSecure Then
		vUrl = "fmodulewindow(\'9082EN|N|N|N\')"
  Else
		vUrl = "window.open(\'/V5/Default.asp?vlang=EN&vCust=VUBZ2274&vId=browser_test&vQModId=9082EN>NN\',\'fModule\',\'toolbar=no,width=750,height=475,left=50,top=50,status=no,scrollbars=no,resizable=no\')"
  End If  
%>

<html>

<head>
  <title>Browser Test</title>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet"> 
  <!-- Check Javascript version -->
  <script type="text/javascript">
    var jsver = 1.0;
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
  <script language="JavaScript">

  function checkOS() {
    if(navigator.userAgent.indexOf('Linux') != -1)
      { var OpSys = "lin"; OpSysDesc = navigator.userAgent.substring(navigator.userAgent.indexOf('Linux'),navigator.userAgent.indexOf(';',navigator.userAgent.indexOf('Linux'))) }
    else if(navigator.userAgent.indexOf('Win') != -1)
      { var OpSys = "win"; OpSysDesc = navigator.userAgent.substring(navigator.userAgent.indexOf('Win'),navigator.userAgent.indexOf(';',navigator.userAgent.indexOf('Win'))) }
    else if(navigator.userAgent.indexOf('Mac') != -1)
      { var OpSys = "mac"; OpSysDesc = navigator.userAgent.substring(navigator.userAgent.indexOf('Mac'),navigator.userAgent.indexOf(';',navigator.userAgent.indexOf('Mac')))}
    else { var OpSys = "other"; OpSysDesc = "Unknown"}

    return OpSys;    
  }
  
  function NamedPair(URL, pair) {   // get the right side of a named pair
    var i = URL.indexOf(pair);      // get the starting position of the named pair
    var j = pair.length;            // get the length of the pair
    var k = i + j;                  // get start of right side
    i = URL.substring(k)            // get the substring
    j = i.indexOf("&")              // next pair starting position?
    if (j == -1)                    // if not then send back full string
      {return (i)} 
    else                            // else just send up to next "&"
      {return (i.substring(0, j))}
  }

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

  function fKillPopup() {
    if (vPopupBlockerWin != null)
      vPopupBlockerWin.close()
  }

  try {
    // Global variables
  
    // Used to extract returnUrl from the incoming URL and determine next page
    var ReturnURL
    var ReturnPage   
    
    // Used for determining Flash
    var ns,ie;
    var platform    = navigator.appVersion.indexOf('Mac') != -1 ? "mac" : "pc";
    var browser     = navigator.appName.indexOf('Netscape') != -1 ? (ns=1) : (ie=1);
    var flashversion= 0;
    var OpSysDesc;
    var btype;
    
    
    // Used to get OS name 
    var OpSys       = checkOS();
  
    beginRollover   = false;  // This handles a bug in Nav4.0x that executes the code too quickly.
  
    // convert all characters to lowercase to simplify testing 
    
    var agt         = navigator.userAgent.toLowerCase(); 
  
    // Browser version *** Note: On IE5, these return 4, so use is_ie5up to detect IE5
     
    var is_major    = parseInt(navigator.appVersion); 
    var is_minor    = parseFloat(navigator.appVersion); 
  
    var is_konq = false;
    var kqPos   = agt.indexOf('konqueror');
    if (kqPos !=-1) {                 
      is_konq  = true;
      is_minor = parseFloat(agt.substring(kqPos+10,agt.indexOf(';',kqPos)));
      is_major = parseInt(is_minor);
    }                                 
  
    var is_safari = ((agt.indexOf('safari')!=-1)&&(agt.indexOf('mac')!=-1))?true:false;
    var is_khtml  = (is_safari || is_konq);
  
    var is_gecko = ((!is_khtml)&&(navigator.product)&&(navigator.product.toLowerCase()=="gecko"))?true:false;
    var is_gver  = 0;
    if (is_gecko) is_gver=navigator.productSub;
  
    var is_fb = ((agt.indexOf('mozilla/5')!=-1) && (agt.indexOf('spoofer')==-1) &&
                 (agt.indexOf('compatible')==-1) && (agt.indexOf('opera')==-1)  &&
                 (agt.indexOf('webtv')==-1) && (agt.indexOf('hotjava')==-1)     &&
                 (is_gecko) && (navigator.vendor=="Firebird"));
    var is_fx = ((agt.indexOf('mozilla/5')!=-1) && (agt.indexOf('spoofer')==-1) &&
                 (agt.indexOf('compatible')==-1) && (agt.indexOf('opera')==-1)  &&
                 (agt.indexOf('webtv')==-1) && (agt.indexOf('hotjava')==-1)     &&
                 (is_gecko) && ((navigator.vendor=="Firefox")||(agt.indexOf('firefox')!=-1)));
    var is_moz   = ((agt.indexOf('mozilla/5')!=-1) && (agt.indexOf('spoofer')==-1) &&
                     (agt.indexOf('compatible')==-1) && (agt.indexOf('opera')==-1)  &&
                    (agt.indexOf('webtv')==-1) && (agt.indexOf('hotjava')==-1)     &&
                    (is_gecko) && (!is_fb) && (!is_fx) &&
                    ((navigator.vendor=="")||(navigator.vendor=="Mozilla")||(navigator.vendor=="Debian")));
    if ((is_moz)||(is_fb)||(is_fx)) {  // 032504 - dmr
       var is_moz_ver = (navigator.vendorSub)?navigator.vendorSub:0;
       if(is_fx&&!is_moz_ver) {
           is_moz_ver = agt.indexOf('firefox/');
           is_moz_ver = agt.substring(is_moz_ver+8);
           is_moz_ver = parseFloat(is_moz_ver);
       }
       if(!(is_moz_ver)) {
           is_moz_ver = agt.indexOf('rv:');
           is_moz_ver = agt.substring(is_moz_ver+3);
           is_paren   = is_moz_ver.indexOf(')');
           is_moz_ver = is_moz_ver.substring(0,is_paren);
       }
       is_minor = is_moz_ver;
       is_major = parseInt(is_moz_ver);
   }
    var is_fb_ver = is_moz_ver;
    var is_fx_ver = is_moz_ver;
  
    var is_nav      = ((agt.indexOf('mozilla')!=-1) && (agt.indexOf('spoofer')==-1) 
                   && (agt.indexOf('compatible') == -1) && (agt.indexOf('opera')==-1) 
                   && (agt.indexOf('webtv')==-1));
    var is_nav2     = (is_nav && (is_major == 2));
    var is_nav3     = (is_nav && (is_major == 3));
    var is_nav4     = (is_nav && (is_major == 4));
    var is_nav6     = (is_nav && (is_major == 5) && !((agt.indexOf('netscape/7')!=-1)||(agt.indexOf('netscape/8')!=-1)));
    var is_nav7     = (is_nav && (agt.indexOf('netscape/7')!=-1));
    var is_nav8     = (is_nav && (agt.indexOf('netscape/8')!=-1));
  
    var is_ie       = (agt.indexOf("msie") != -1); 
    var is_ie3      = (is_ie && (is_major < 4)); 
    var is_ie4      = (is_ie && (is_major == 4) && (agt.indexOf("msie 4")!=-1) ); 
    var is_ie5      = (is_ie && (is_major == 4) && (agt.indexOf("msie 5")!=-1) ); 
    var is_ie6      = (is_ie && (is_major == 4) && (agt.indexOf("msie 6")!=-1) ); 
    var is_ie7      = (is_ie && (is_major == 4) && (agt.indexOf("msie 7")!=-1) ); 
    var is_ie8      = (is_ie && (is_major == 4) && (agt.indexOf("msie 8")!=-1) ); 
    var is_ie9      = (is_ie && (is_major == 4) && (agt.indexOf("msie 9")!=-1) ); 
  
    // IE
    if (is_ie3) {
    	btype = 'ie';
    	bver = '3'
    }
    else if (is_ie4) {
    	btype = 'ie';
    	bver = '4'
    }
    else if (is_ie5) {
    	btype = 'ie';
    	bver = '5'
    }
    else if (is_ie6) {
    	btype = 'ie';
    	bver = '6'
    }
    else if (is_ie7) {
    	btype = 'ie';
    	bver = '7'
    }
    else if (is_ie8) {
    	btype = 'ie';
    	bver = '8'
    }
    else if (is_ie9) {
    	btype = 'ie';
    	bver = '9'
    }
    // Netscape
    else if (is_nav3) {
    	btype = 'ns';
    	bver = '3'
    }
    else if (is_nav4) {
    	btype = 'ns';
    	bver = '4'
    }
    else if (is_nav6 && !is_nav7) {
    	btype = 'ns';
    	bver = '6'
    }
    else if (is_nav7) {
    	btype = 'ns';
    	bver = '7'
    }
    else if (is_nav8) {
    	btype = 'ns';
    	bver = '8'
    }
    // Firefox
    else if (is_fx) {
    	btype = 'fx';
    	bver = is_fx_ver
    }
  
    // Comprehensive Flash detection
  
    theVBScript =  '<form name="vbform"><input type="hidden" name="flashdetect"><input type="hidden" name="rpdetect">\</form>\n';
    theVBScript += '<SCR' + 'IPT LANGUAGE="VBScript">\n';
    theVBScript += 'on error resume next\n';
    theVBScript += 'RealPyr = "False"\n';
    theVBScript += 'Set oTest = CreateObject("ShockwaveFlash.ShockwaveFlash.2")\n';
    theVBScript += 'If Err = 0 Then FlashInstalled = 2\n';
    theVBScript += 'Err.Clear\n';
    theVBScript += 'Set oTest = CreateObject("ShockwaveFlash.ShockwaveFlash.3")\n';
    theVBScript += 'If Err = 0 Then FlashInstalled = 3\n';
    theVBScript += 'Err.Clear\n';
    theVBScript += 'Set oTest = CreateObject("ShockwaveFlash.ShockwaveFlash.4")\n';
    theVBScript += 'If Err = 0 Then FlashInstalled = 4\n';
    theVBScript += 'Err.Clear\n';
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
  
    if (ns || is_fx){
        var FlashInstalled = 0;
        nparray = navigator.plugins;
        nparraylen = nparray.length;
        RealPyr = 'False';
  
        // Search for Flash
  
        for (i=0;i<nparraylen;i++){
    		npplugin = nparray[i];
    		npname   = npplugin.name;
    		npdesc   = npplugin.description;
  
    		if (npname.indexOf("Shockwave Flash 2.") != -1) { flashversion = 2; break }
    		if (npdesc.indexOf("Shockwave Flash 3.") != -1) { flashversion = 3; break }
    		if (npdesc.indexOf("Shockwave Flash 4.") != -1) { flashversion = 4; break }
    		if (npdesc.indexOf("Shockwave Flash 5.") != -1) { flashversion = 5; break }
    		if (npdesc.indexOf("Shockwave Flash 6.") != -1) { flashversion = 6; break }
    		if (npdesc.indexOf("Shockwave Flash 7.") != -1) { flashversion = 7; break }
    		if (npdesc.indexOf("Shockwave Flash 8.") != -1) { flashversion = 8; break }
    		if (npdesc.indexOf("Shockwave Flash 9.") != -1) { flashversion = 9; break }
    		if (npdesc.indexOf("RealPlayer") != -1) { RealPyr = 'True'}
  
        }
        if (flashversion>1)
          FlashInstalled = flashversion
    
    // Search for RealPlayer
  
    } else if (ie){
    	document.write(theVBScript);
    	flashversion = document.vbform.flashdetect.value;
    	RealPyr = document.vbform.rpdetect.value;
    }
  
    // Added by Mike to allow for IE on MAC
    if(ie && OpSys == 'mac' ){
      flashversion = 0;
    }
  
    // get the Time Zone offset (GMT)
    var ZoneOffset;
    ClientDate = new Date();
    ZoneOffset = ClientDate.getTimezoneOffset()/60;
  
    // Check if popup blocker is installed...if it opens...we will close it when we unload
    var vPopupBlockerWin = window.open('Popup.htm','','left=2000,top=2000,width=10,height=10')
    if (vPopupBlockerWin==null)
      vPopupBlocker=true
    else {
      vPopupBlocker=false
    }
  
    // Check if user allows Cookies
    var vCookies = true
    createCookie('VUTestCookie','written',10)
    if(readCookie('VUTestCookie')==null) vCookies = false
    eraseCookie('VUTestCookie')
  
    var vHTML
    
    vHTML  = '		<P align="center">'
    vHTML += '			<TABLE class=c2 style="border-collapse: collapse" bordercolor="#DDEEF9"" cellSpacing="0" cellPadding="10" width="100%" border="1" ID="Table1" align="right">'
    vHTML += '				<TR>'
    vHTML += '					<TD><B>Component</B></TD>'
    vHTML += '					<TD><B>Details</B></TD>'
    vHTML += '					<TD><B>Status</B></TD>'
    vHTML += '					<TD><B>Notes</B></TD>'
    vHTML += '				</TR>'
    vHTML += '				<TR>'
    vHTML += '					<TD>Operating System</TD>'
    vHTML += '					<TD>' + OpSysDesc + '</TD>'
    vHTML += '					<TD>n/a</TD>'
    vHTML += '					<TD>&nbsp;</TD>'
    vHTML += '				</TR>'
    vHTML += '				<TR>'
    vHTML += '					<TD>Javascript Version</TD>'
    vHTML += '					<TD>' + jsver + '</TD>'
    vHTML += '					<TD>n/a</TD>'
    vHTML += '					<TD>&nbsp;</TD>'
    vHTML += '				</TR>'
    vHTML += '				<TR>'
    vHTML += '					<TD>Screen Resolution</TD>'
    vHTML += '					<TD>' + window.screen.width + ' x  ' + window.screen.height + ' pixels</TD>'
    vHTML += '					<TD>n/a</TD>'
    vHTML += '					<TD>&nbsp;</TD>'
    vHTML += '				</TR>'
    vHTML += '				<TR>'
    vHTML += '					<TD>Colour Depth</TD>'
    vHTML += '					<TD>' + window.screen.colorDepth + ' bit</TD>'
    vHTML += '					<TD>n/a</TD>'
    vHTML += '					<TD>&nbsp;</TD>'
    vHTML += '				</TR>'
    vHTML += '				<TR>'
    vHTML += '					<TD>Browser</TD>'
    if (btype=='ie')
      vHTML += '					<TD>Internet Explorer ' + bver + '</TD>'
    else if (btype=='ns')
      vHTML += '					<TD>Netscape ' + bver + '</TD>'
    else if (btype=='fx')
      vHTML += '					<TD>Firefox ' + bver + '</TD>'
    else
      vHTML += '					<TD>Not known</TD>'
    // if ((btype=='ie'&&bver>=6)||(btype=='ns'&&bver>=7)||(is_fx)) {
    // only allow IE
    if (btype=='ie'&&bver>=6) {
      vHTML += '					<TD><FONT face="Verdana" color="green">Passed</FONT></TD>'
      vHTML += '					<TD>&nbsp;</TD>'
    }
    else if (btype=='fx') {
      vHTML += '					<TD><FONT face="Verdana" color="green">Passed</FONT></TD>'
      vHTML += '					<TD>&nbsp;</TD>'
    }
    else {
      vHTML += '					<TD><FONT face="Verdana" color="red">Warning</FONT></TD>'
      vHTML += '					<TD>Only then IE 5.5+ and Firefox browsers are accepted.  PC users can update your browser...<A href="http://www.microsoft.com/windows/ie" target="_blank">www.microsoft.com/windows/ie</A></TD>'
    }
    vHTML += '				</TR>'
    vHTML += '				<TR>'
    vHTML += '					<TD>Pop-up Blocker</TD>'
    if (vPopupBlocker) {
      vHTML += '					<TD>yes</TD>'
      vHTML += '					<TD><FONT face="Verdana" color="red">Failed</FONT></TD>'
      vHTML += '					<TD>You must allow pop-ups.</TD>'
    }
    else {
      vHTML += '					<TD>no</TD>'
      vHTML += '					<TD><FONT face="Verdana" color="green">Passed</FONT></TD>'
      vHTML += '					<TD>&nbsp;</TD>'
    }
    vHTML += '				</TR>'
    vHTML += '				<TR>'
    vHTML += '					<TD>Cookies</TD>'
    if (vCookies) {
      vHTML += '					<TD>allowed</TD>'
      vHTML += '					<TD><FONT face="Verdana" color="green">Passed</FONT></TD>'
      vHTML += '					<TD>&nbsp;</TD>'
    }
    else {
      vHTML += '					<TD>not allowed</TD>'
      vHTML += '					<TD><FONT face="Verdana" color="red">Failed</FONT></TD>'
      vHTML += '					<TD>Ensure that the option to store cookies is turned on within your browser options.</TD>'
    }
    vHTML += '				</TR>'
    vHTML += '				<TR>'
    vHTML += '					<TD>Flash</TD>'
    if (FlashInstalled>1) {
      vHTML += '					<TD>Flash ' + flashversion + '</TD>'
      vHTML += '					<TD><FONT face="Verdana" color="green">Passed</FONT></TD>'
      vHTML += '					<TD>&nbsp;</TD>'
    }
    else {
      vHTML += '					<TD>not installed</TD>'
      vHTML += '					<TD><FONT face="Verdana" color="red">Failed</FONT></TD>'
      vHTML += '					<TD>Flash must be installed. Download at <A href="http://www.macromedia.com" target="_blank">www.macromedia.com</A>.</TD>'
    }
    vHTML += '				</TR>'
    vHTML += '				<TR>'
    vHTML += '					<TD>If Passed then click...</TD>'
    if (!vPopupBlocker) {
      vHTML += '					<TD><input type="button" onclick="<%=vUrl%>" value="Launch Module" name="bLaunch" class="button"></TD>'
      vHTML += '					<TD>n/a</TD>'
      vHTML += '					<TD>&nbsp;</TD>'
    }
    else {
      vHTML += '					<TD>Unavailable (Popup Blocker present)</TD>'
      vHTML += '					<TD>n/a</TD>'
      vHTML += '					<TD>This will become active once your Popup Blocker issue is resolved.</TD>'
    }
    vHTML += '				</TR>'
    vHTML += '				<TR>'
    vHTML += '					<TD colspan="4" align="center"><br><br><input onclick="window.close()" type="button" value="Close" name="bClose" class="button"><br><br></TD>'
    vHTML += '				</TR>'
    vHTML += '			</TABLE>'
    vHTML += '		</P>'
  
  
    //QString = 'zoneoffset=' + ZoneOffset + '\nplatform=' + OpSys + '\nbrowser=' + btype + '\nbver=' + bver + '\nflashver=' + flashversion + '\npopup blocker=' + vPopupBlocker + '\nallow cookies=' + vCookies;
    //alert(QString)
  }
  catch (e) {
    alert('An error has been encountered.  The following are the error details\n  Number: ' + e.number + '\n  Description: ' + e.description)
  }
   
  
  // use this rather then the version in Functions.js since we don't trackwindows, etc
  function fmodulewindow(vModId)
  {
    var newTop  = (screen.height - 475) / 2 
    var newLeft = (screen.width  - 750) / 2    
    if (newTop  < 1) {newTop  = 0}
    if (newLeft < 1) {newLeft = 0}
    var vmodule = "/V5/LaunchObjects.asp?vModId=" + vModId
    var modwindow = window.open(vmodule,'Module','toolbar=no,width=750,height=475,left='+newLeft+',top='+newTop+',status=no,scrollbars=no,resizable=yes')
//  top.addWindowToArray(modwindow) 
    modwindow.focus()
    parent.vModWindow = modwindow
    parent.vModWindowOpen = true
  }


  </script>
</head>

<body onunload="fKillPopup()" bgcolor="#FFFFFF">

  <table width="100%" border="0" cellspacing="0" cellpadding="10">
    <tr>
      <td width="100%" align="center">
      <h1 align="center">Browser Test</h1>
        <p class="c2" align="left">If the Browser, Popup Blocker, Cookies and Flash all read <font face="Verdana" color="green">Passed</font>, then you should be able to launch a test module by clicking the <b>Launch Module</b> button below.</p>
        <p>
        <script>
          document.write(vHTML)
        </script>
        </td>
    </tr>
  </table>

  </body>

</html>