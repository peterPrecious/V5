<!--
  This routine is strictly used when a module is closed in popup free mode.
  It assumes the module is being rendered in a Div called "Div_Module".
  When this page is called we want to close that div.
-->

<html>

<head>
  <meta http-equiv="Content-Language" content="en-us">
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <title>Black Page</title>
  <script>

    function fCloseDiv() {
      try {
        parent.parent.iModule.parent.document.getElementById('Div_Module').style.display='none';
      }
      catch(err) {
//      alert("no parent.parent");
        try {
          parent.iModule.parent.document.getElementById('Div_Module').style.display='none';
        }
        catch(err) {
//        alert("no parent");
        }
      }
    }
  </script>
</head>

<body onload="fCloseDiv()" bgcolor="#000000"></body>

</html>