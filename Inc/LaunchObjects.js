

  //   the older launch window functions are in the "Launch.js" - this version is generic for modules containing X,Y values
  function launchObjects (objId, objWidth, objHeight) {
    var url = "../LaunchObjects.asp?vModId=" + objId
    var modwindow = window.open(url,'object','width=' + objWidth + ',height=' + objHeight + ',toolbar=no,left=10,top=10,status=no,resizable=yes,scrollbars=yes')
    top.addWindowToArray(modwindow) 
    modwindow.focus()
    parent.vModWindow = modwindow
    parent.vModWindowOpen = true
  }