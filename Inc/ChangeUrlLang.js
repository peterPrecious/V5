    function ChangeURLLang(vNewLang) 
    {
      var vUrl
      if (location.search.length>0)
        if (location.href.toLowerCase().indexOf('vlang=')>0)
          vUrl=location.href.replace(location.href.substring(location.href.toLowerCase().indexOf('vlang='),location.href.toLowerCase().indexOf('vlang=')+8),'vLang=' + vNewLang)
        else
          if (location.href.indexOf('\#')>0)
            vUrl=location.href.substring(0,location.href.indexOf('\#')) + '&vLang=' + vNewLang + location.href.substring(location.href.indexOf('\#'))
          else
            vUrl=location.href + '&vLang=' + vNewLang
      else
        vUrl=location.href + '?vLang=' + vNewLang
      return vUrl
    }
    