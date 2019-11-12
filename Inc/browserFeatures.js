function browserFeatures() {

  var url, browser, flash, html5, touch, cookies, popups, dummy;

  try {
    html5 = Modernizr.canvas ? "y" : "n";
    touch = Modernizr.touch ? "y" : "n";
  } catch (e) {
    html5 = "n";
    touch = "n";
  }

  try {
    var agent = navigator.userAgent.toLowerCase();
    if (agent.indexOf("msie") > 0) { browser = "IE" }
    else if (agent.indexOf("trident/") > 0) { browser = "IE 11" }
    else if (agent.indexOf("edge/") > 0) { browser = "MS Edge" }
    else if (agent.indexOf("firefox/") > 0) { browser = "Firefox" }
    else if (agent.indexOf("opr/") > 0) { browser = "Opera" }
    else if (agent.indexOf("chrome/") > 0) { browser = "Chrome" }
    else if (agent.indexOf("safari/") > 0) { browser = "Safari" }
    else if ($.browser.msie) { browser = "MSIE" }
    else if ($.browser.safari) { browser = "Safari" }
    else if ($.browser.mozilla) { browser = "Mozilla" }
    else if ($.browser.opera) { browser = "Opera" }
    else if ($.browser.webkit) { browser = "Webkit" }
    else { browser = "other" }

    //just add next line for the /popups dev test
    //browser = "<a href='#' onclick='$(\"#version\").toggle()'>" + browser + "</a> <span id='version' style='display:none'>(" + agent + ")</span>";

  } catch (e) {
    browser = "n/a";
  }

  // this is a temp hack to skip popup test if edge
  var edge = (agent.indexOf("edge/") == -1) ? false : true;

  try {
    flash = swfobject.getFlashPlayerVersion().major;
  } catch (err) {
    flash = "n/a"
  }
  flash = flash == "0" ? "n/a" : flash;

  cookies = (navigator.cookieEnabled) ? "y" : "n";

  // if enabled (ie NO blocker) show y
  popups = "y";
  if (!edge) {
    try {
      dummy = window.open('Inc/popupBlocker.htm', '', 'width=100,height=100,left=10,top=10,scrollbars=no,location=no,menubar=no,toolbar=no,statusbar=no');
      if (dummy) {
        popups = "y";
        dummy.close();
      }
      else {
        popups = "n";
      }
    }
    catch (e) {
      popups = "n";
    }
  }

  // note: ecomOk is derived in Default.asp head section
  // alert(touch + "|" + browser + "|" + html5 + "|" + flash + "|" + cookies + "|" + popups + "|" + ecomOk + "|" + navigator.userAgent.toLowerCase());
  return (touch + "|" + browser + "|" + html5 + "|" + flash + "|" + cookies + "|" + popups + "|" + ecomOk + "|" + navigator.userAgent.toLowerCase());
}