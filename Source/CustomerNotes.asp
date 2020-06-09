<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
	Server.Execute vShellHi
  %>
  <div style="width: 400px; margin: auto;">
    <h1 align="center">Customer Notes</h1>
    <p class="c2">This creates a spreadsheet of Accounts containing customer Notes.&nbsp; <br>To be included, at least one of the 5 Note fields must contain data.</p>
    <p class="c2">
    <input onclick="location.href = 'CustomerNotes_x.asp'" type="button" value="Go" name="bGo" class="button"></p>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
