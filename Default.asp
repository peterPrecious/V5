<%@Language=VBScript CodePage = 65001%>

<%
  '... these top 9 lines were added Nov 12, 2019 because certain browsers (IE) were no longer handling accents, oddly Chrome was OK
  Session.CodePage = 65001
  Response.charset ="utf-8"
  Session.LCID     = 1033 'en-US
%>

<html>

<head>
  <title>V5 Redirector</title>
  <script src="/V5/Inc/jQuery.js"></script>
  <script src="/V5/Inc/Functions.js"></script>

  <!-- check TLS status script (run even if not needed)-->
  <script>
  var ecomOk = "n";
  window.parseTLSinfo = function (data) {
    var version = data.tls_version.split(' ');
    ecomOk = version[0] != 'TLS' || version[1] < 1.2 ? "n" : "y";
  }
  </script>
  <script src="https://www.howsmyssl.com/a/check?callback=parseTLSinfo"></script>

  <script src="/V5/Inc/browserFeatures.js"></script>
  <script src="/V5/Scripts/modernizr.js"></script>
  <script src="/V5/Scripts/swfobject.js"></script>
</head>

<body>

  <!-- this form is a hack to redirect an IBAO V5 request to V8 using a form post  -->
  <!--
  <form id="ibao" action="/vubizApps/IBAOsignUp.aspx" method="post" target="_self" style="display: none;">
  -->
  <form id="ibao" action="/vubizApps/IBAOtoNOP.aspx" method="post" target="_self" style="display: none;">
    <input type="hidden" id="membId" name="membId" value="" />
  </form>

  <script>
    $(function () {

      /* IBAO hack - legacy allows them to stay in old V5 */
      if (getParameter("vCust") == "NSRC2321" && getParameter("vAction") != "legacy") { 
          $("#membId")[0].value = getParameter("vID");
          $("#ibao").submit();

      /* Normal V5 Access */
      } else {   
        var url = "Start.asp?vVer=13";
        /* grab the post values from server and/or the querystring values from URL */
        var formItems = "<%=Request.Body.Item%>";
        var queryString = location.search.replace("?","");
        if (queryString != "")  { 
          url = url + "&" + queryString;
        }
        else if (formItems != "") { 
	        url = url + "&" + formItems; 
        }; 
        url = url + "&vBrowser=" + browserFeatures();
        window.location.assign(url);						      
      }
    })

  </script>
</body>
</html>
