<script>

  function getParameter(name) {
    // Parses the querystring
    var pair=location.search.substring(1).split("&");
  
    for (var i=0;i<pair.length;i++) {
      var a=pair[i].split("=");
      var n="",v="";
      if (a.length>0)
        n=unescape(a[0]);
      if (a.length>1)
        v=unescape(a[1]);
      if (n.toLowerCase()==name.toLowerCase()) return v;
    }
    return null;       
  }

  location.href = getParameter('vNext') + location.search + '#' + getParameter('vAnchor')

</script>
