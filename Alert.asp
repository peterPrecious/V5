<%
  If Request("vPwd") = "21122112" Then
    Application("Alert") = Request("vAlert") 
  End If
%>

<html>

<head>
  <meta charset="UTF-8">
  <title>:: VuAlert</title>
</head>

<body>
      <form method="POST" action="Alert.asp">
        <p>Put Alert: <input type="radio" value="y" name="vAlert">On<input type="radio" value="n" name="vAlert" checked>Off </p>
        <p>Password: <input type="text" name="vPwd" size="9"></p>
        <p><input type="submit" value="Submit" name="bAlert"></p>
      </form>
</body>

</html>
