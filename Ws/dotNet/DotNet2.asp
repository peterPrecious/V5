<%
  
  
  
  
  %>


<!DOCTYPE html>
<html>
<head>
  <title>Calling a C# webservice from classic ASP & Javascript</title>
  <script src="//code.jquery.com/jquery-1.11.2.min.js"></script>
  <meta charset="UTF-8">
</head>
<body>

  <%
    If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
      Dim xmlhttp
      Dim DataToSend
'     DataToSend="val1="&Request.Form("text1")&"&val2="&Request.Form("text2")
      DataToSend="plainText=lâcher un pet"
      Dim postUrl
'     postUrl = "http://localhost/DummyWS/WebService1.asmx/testDrivel"
      postUrl = "http://localhost/DummyWS/WebService1.asmx/Base64Encode"
      Set xmlhttp = server.Createobject("MSXML2.XMLHTTP")
      xmlhttp.Open "POST", postUrl, false
      xmlhttp.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
      xmlhttp.send DataToSend
      Response.Write(xmlhttp.responseText)
    End If
  %>

  <form method="POST" name="form1" id="Form1">
    <input type="text" name="text1" id="Text1" value="12"><br />
    <input type="text" name="text2" id="Text2" value="23"><br />
    <input type="submit" value="GO" name="submit1" id="Submit1">
  </form>


</body>

<!--  
  <script>
    $.ajax( {
      type:       "POST",
      data:       { 'val1': '12', 'val2': '23' },
      url:        "http://localhost/DummyWS/WebService1.asmx/fu",
      success:    function(data) { alert(data.all[0].innerHTML); },
      error:      function (jqXHR, textStatus, errorThrown) { alert(jqXHR.responseText); }
    })
  </script>

  <script>
    $.ajax( {
      type:       "POST",
      data:       { 'plainText': 'lâcher un pet' },
      url:        "http://localhost/DummyWS/WebService1.asmx/Base64Encode",
      success:    function(data) { alert(data.all[0].innerHTML); },
      error:      function (jqXHR, textStatus, errorThrown) { alert(jqXHR.responseText); }
    })
  </script>
-->


</html>
