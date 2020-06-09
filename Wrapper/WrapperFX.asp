<html>
  <head>
    <meta charset="UTF-8">
    <script>
      function getParameter(name) {
        var pair=location.search.substring(1).split("&");    
        for (var i = 0; i < pair.length; i++) {
          var a = pair[i].split("=");
          var n="",v="";
          if (a.length > 0)
            n = unescape(a[0]);
          if (a.length > 1)
            v = unescape(a[1]);
          if (n.toLowerCase() == name.toLowerCase()) return v;
        }
        return null;       
      } 
      function SCORM_Finish() {
        return null;
      }
      var vModId  = getParameter("vModId")
      var vTitle  = getParameter("vTitle")
      var vToken  = getParameter("vToken")
      var vUrl = "/V5/fModules/" + vModId + "/lmsstart.html";
alert(vUrl);


    </script>
    <script language="JavaScript" src="/V5/Wrapper/indexScorm.js"></script>
    <script>
      document.write   ('    <title>' + getParameter("vTitle") + '</title>');
      document.write   ('  </head>');
      document.write   ('  <frameset rows="100%" onbeforeunload="SCORM_Finish()" onunload="SCORM_Finish()">');
      document.write   ('    <frame name="content" src="' + vUrl  + '" frameborder="no" scrolling="no" noresize marginwidth="0" marginheight="0">');
      document.write   ('  </frameset>');
    </script>
</html>
