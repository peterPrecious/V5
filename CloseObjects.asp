
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
  <head>
    <script src="/V5/Inc/jQuery.js"></script>
    <script type="text/javascript">
      function closeObjects() {
        // if exists, pass URL parameters to the logging code
        if (location.search.length == 0) {
          var url = "/V5/Logx.asp?vSource=CloseObjects";
        } else {
          var url = "/V5/Logx.asp" + location.search + "&vTerminate=y&vSource=CloseObjects";
        }
//      alert(url);
        $.post(url);
        window.open('', '_parent', '');
        window.close();
      }
    </script>
  </head>
  <body onload="closeObjects()"></body>
</html>
