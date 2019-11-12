if (document.layers) {
  document.captureEvents(Event.MOUSEDOWN);
  document.onmousedown = clickNS4;
}
else if (document.all && !document.getElementById) {
  document.onmousedown = clickIE4;
}

//  document.oncontextmenu=new Function("return false")


function fModule(vMod) {
  modwindow = window.open('//author.vubiz.com/fModules/' + vMod + '/default.htm?vModId=' + vMod, '', 'resizable=yes,width=750,height=475');
}

function flex(vMod) {
  modwindow = window.open('//author.vubiz.com/fModules/' + vMod + '/LMSStart.html?vModId=' + vMod + '&deployment=vubuild', '', 'resizable=yes,width=775,height=570');
}

function HTML(vMod) {
  modwindow = window.open('//author.vubiz.com/fModules/' + vMod + '/DefaultHTML.htm?vModId=' + vMod + '&deployment=vubuild', '', 'resizable=yes,width=800,height=600');
}
/*
    function clickIE4()
    {
      if (event.button==2)
      {
        return false;
      }
    }

    function clickNS4(e)
    {
      if (document.layers||document.getElementById&&!document.all)
      {
        if (e.which==2||e.which==3)
        {
          return false;
        }
      }
    }    

*/