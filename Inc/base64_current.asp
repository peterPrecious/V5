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