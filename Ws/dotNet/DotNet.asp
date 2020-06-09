<!--#include virtual = "ChAccess\Inc\InitializeCH.asp"-->

<html>

<head>
  <title>IBEW Landing Page</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/WebService.js"></script>
  <script src="/V5/Inc/jQuery.js"></script>
  <script src="wsDotNet.js"></script>
  <script>
    function validate()
    {
      var ws, param
    
      param = 'inp=' + $("#inp").val();
      ws = WebService('wsDotNet.asp', param)
      $("#out").val(ws);

    }
  </script>
</head>

<body>


  <input id="inp" type="text" value="12" />
  <input id="out" type="text" value="0" />
  <input id="Button1" onclick="validate()" type="button" value="button" />
</body>

</html>
