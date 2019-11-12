
  // these two functions have been incorporated into Functions.js
  // this file is called by the assessment player

  function WebService(vUrl, vMsg) {

    var agt    = navigator.userAgent.toLowerCase(); 
    var ie     = (agt.indexOf("msie") != -1); 
    if (ie)
      oXmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
    else
      oXmlHttp = new XMLHttpRequest();

    try {
      oXmlHttp.open("POST", vUrl, false);
      oXmlHttp.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
      oXmlHttp.send(vMsg);    
      return oXmlHttp.responseText;

    }
    catch (err) {
      alert(err);
      return "error using web service";
    }  

  }
  
  
  
  function jsonWebService(vUrl, vMsg) {

    var agt    = navigator.userAgent.toLowerCase(); 
    var ie     = (agt.indexOf("msie") != -1); 
    if (ie)
      oXmlHttp = new ActiveXObject("Microsoft.XMLHTTP");
    else
      oXmlHttp = new XMLHttpRequest();

    try {
      oXmlHttp.open("POST", vUrl, false);
      oXmlHttp.setRequestHeader("Content-Type", "application/json; charset=utf-8");
      oXmlHttp.send(vMsg);    
      return oXmlHttp.responseText;

    }
    catch (err) {
      alert(err);
      return "error using web service";
    }  

  }  