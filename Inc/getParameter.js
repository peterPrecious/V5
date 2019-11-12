   //  this returns a querystring value and strips off any "+" representing spaces

    function getParameter(name) {
      var pair=location.search.substring(1).split("&");
      for (var i = 0; i < pair.length; i++) {
        var a = pair[i].split("=");
        var i, n="", v="";
        if (a.length > 0) {
          n=a[0];
          if (n==name) {
            if (a.length > 1) {
              v=unescape(a[1]);
              for (i = 0;  i < v.length; i++) {
                v = v.replace('+', ' ');
              }  
              return v;               
            }  
          }
        }
      }
    }
