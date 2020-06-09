<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>Accent handling: Server.URLEncode(plainText) </title>
</head>
<body>
  <%
   
    Dim plainText    
    plainText = "&vFirstName=Péter&vLastName=Bulloch&vScore=80&vDate=Jan 29, 2020&vModsId=1234EN&vTitle=Test Assessment&vLang=FR&vCust=VUBZ&vAcctId=2274&vProgId=P1234&vLogo=vubz.jpg&vMemo=||0.83|Péter: Eat My Shorts|&vEmailTo=1234567&vEmailFrom=info@vubiz.com"

    Dim xmlhttp, dataToSend, postUrl
    dataToSend = "plainText=" & Server.URLEncode(plainText)
    postUrl = "http://" & Request.ServerVariables("HTTP_HOST") & "/V5/Inc/base64.asmx/base64Encode"
    Set xmlhttp = server.Createobject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", postUrl, false
    xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
    xmlhttp.send dataToSend
    base64Encode = xmlhttp.responseXML.text
    
    
    encodedText = base64Encode

    dataToSend = "base64EncodedData=" & encodedText
    postUrl = "http://" & Request.ServerVariables("HTTP_HOST") & "/V5/Inc/base64.asmx/base64Decode"
    Set xmlhttp = server.Createobject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", postUrl, false
    xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
    xmlhttp.send dataToSend
    base64Decode = xmlhttp.responseXML.text

    stop
   
    
    
    %>
</body>
</html>