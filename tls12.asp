<!DOCTYPE html>
<html>
<head>
  <title></title>
  <meta charset="utf-8" />

</head>
<body>
  <%
    Set objHttp = Server.CreateObject("WinHTTP.WinHTTPRequest.5.1")
'   objHttp.open "GET", "http://howsmyssl.com/a/check", False
    objHttp.open "GET", "https://www.ssllabs.com/ssltest/viewMyClient.html", False

    objHttp.Send
    Response.Write objHttp.responseText 
    Set objHttp = Nothing 




  %>
</body>
</html>
