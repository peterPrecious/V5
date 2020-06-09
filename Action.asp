<html>
	<head>
		<meta charset="UTF-8">
		<title>Action</title>
	  <script type="text/javascript" src="/V5/Inc/Functions.js"></script>
	  <script type="text/javascript">
	    function action() {
	      var vWs    = WebService("Action_ws.asp", "vAction=" + getParameter("vAction") + "&vCust=" + getParameter("vCust") + "&vLearnerId=" + getParameter("vLearnerId") + "&vManagerId=" + getParameter("vManagerId"));
	      window.open('', '_parent', '');
	      window.close();
	    }
	  </script>
	</head>
  <body onload="action()"></body>
</html>
