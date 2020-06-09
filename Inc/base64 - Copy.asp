


<%
  '...functions copied from certificate.asp
  Function base64Encode(plainText)
    Dim xmlhttp, dataToSend, postUrl
    dataToSend = "plainText=" & plainText
    postUrl = "http://" & Request.ServerVariables("HTTP_HOST") & "/V5/Inc/base64.asmx/base64Encode"
    Set xmlhttp = server.Createobject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", postUrl, false
    xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
    xmlhttp.send dataToSend
    base64Encode = xmlhttp.responseXML.text
  End Function

  Function base64Decode(encodedText)
    Dim xmlhttp, dataToSend, postUrl
    dataToSend = "base64EncodedData=" & encodedText
    postUrl = "http://" & Request.ServerVariables("HTTP_HOST") & "/V5/Inc/base64.asmx/base64Decode"
    Set xmlhttp = server.Createobject("MSXML2.XMLHTTP")
    xmlhttp.Open "POST", postUrl, false
    xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
    xmlhttp.send dataToSend
    base64Decode = xmlhttp.responseXML.text
  End Function

   
%>

<!DOCTYPE html>
<html>
<head>
  <title>Calling a C# webservice from classic ASP</title>
  <meta charset="utf-8" />
</head>
<body>

</body>

</html>
  <%
    Dim plainText, encodedText, decodedText

'    plainText = "These contain accents: lâcher and déshabiller."
    plainText = "&vFirstName=Péter&vLastName=Bulloch&vScore=80&vDate=Jan 28, 2020&vModsId=1234EN&vTitle=Test Assessment&vLang=FR&vCust=VUBZ&vAcctId=2274&vProgId=P1234&vLogo=vubz.jpg&vMemo=||0.83|Péter: Eat My Shorts|&vEmailTo=1234567&vEmailFrom=info@vubiz.com"
'    plainText = "&vFirstName=Péter&vLastName=Bulloch"

'   plainText = Replace(plainText, "&", "&amp;")
    plainText = Replace(plainText, "&", "||")


    response.Write ("<br>encoding : " & plainText)

    encodedText = base64Encode(plainText)   
    response.Write ("<br>encoded&nbsp; : " & encodedText)

    decodedText = base64Decode(encodedText)   
    response.Write ("<br>decoded&nbsp; : " & decodedText)


  %>