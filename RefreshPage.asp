<html>
<head>
	<script src="/V5/Inc/jQuery.js"></script>
  <script>  
    var bodyFocus = true;      // start with focus ON
    function launch() {
      popupWindow = window.open('','','width=10, height=10');
      popupWindow.focus();
      bodyFocus = false;
    }   
    $(function() {
			$(window).focus (
				function() {
		      if (!bodyFocus) {
			      bodyFocus = true;
						location.reload();
			    }
				}
			);
		}	);
  </script>
</head>

<body>
  <div id="content">
    <% if len(session("count")) > 0 Then session("count") = session("count") + 1 else session("count") = 1 %> 
    <% response.write "Refresh: " & session("count") %> 
    <a href="javascript:launch()">Launch</a><br>
  </div>
</body>

</html>
