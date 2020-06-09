<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<% Session.TimeOut = 1 %>

<html>

<head>
  <meta charset="UTF-8">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <script src="/V5/Inc/Functions.js"></script>
</head>

<body>

  <div class="div" id="divBackground" style="height: 100%; width: 100%; position:absolute; left:1px; top:-1px; background-color:#FFFFFF" style="position: absolute;">
    <div class="div" id="sessionAlert" name="sessionAlert" align="center" style="position: absolute; background-color:#ffffff">
      <table cellpadding="10" style="border-collapse: collapse" border="1" bordercolor="#FF0000" width="100%">
        <tr>
          <td class="c5" align="center">
            <!--webbot bot='PurpleText' PREVIEW='Your session will expire in'--><%=fPhra(000979)%> <span id="dTimeRemaining"></span>.<br><br>
            <!--webbot bot='PurpleText' PREVIEW='If you need more time click here.'--><%=fPhra(000980)%><br><br>
            <input onclick="sessionRestart();" type="button" value="<%=bContinue%>" name="bContinue" class="button"> 
          </td>
        </tr>
      </table>
    </div>
    <div class="div" id="sessionEnd" name="sessionEnd" align="center" style="position: absolute; background-color:#ffffff">
      <table cellpadding="10" style="border-collapse: collapse" border="1" bordercolor="#FF0000" width="100%">
        <tr>
          <td class="c5" align="center">
            <!--webbot bot='PurpleText' PREVIEW='Your session has expired.'--><%=fPhra(000981)%><br><br>
            <!--webbot bot='PurpleText' PREVIEW='Please sign-into this service again to resume your learning.'--><%=fPhra(000200)%>
          </td>
        </tr>
      </table>
    </div>
  </div>

  <div id="pooh"></div>

  <script>   
    var HH, MM, SS, timerId, timeRemaining, timeStart, timeComplete, showDiv;
   
    sessionStart();
    sessionStatus(); 

    // this is the duration of the session in hours, minutes and seconds
    function sessionStart() {       
      HH = 0;
      MM = <%=Session.TimeOut%>;
      SS = 0;
      showDiv = false;
      timerId = null;
    }

    // this is run every 10 seconds and the time is decreased by 1 secs
    function sessionStatus() {       

      //   when we reach this stage alert user
      if ((SS == 30) && (MM == 0) && (HH == 0)) {
        showAlert('sessionAlert');       
        showDiv = true;
      }  

      //   when we reach this stage the session has expired
      if ((SS == 0) && (MM == 0) && (HH == 0)) {
        divOff('sessionAlert');
        showAlert('sessionEnd');       
        clearTimeout(timerId);
        return "end";
      }  

      //   start counting down 
      if (!((SS == 0) && (MM == 0) && (HH == 0))) {
        if ((MM == 0) && (SS == 0)) {
          MM = 59;
          SS = 59;
          HH = HH - 1;
        }
        else if (SS == 00) {
          SS = 59;
          MM = MM - 1;
        }
        else {
          SS = SS - 1;
        }
      }

      timeRemaining  = "";
      if (HH > 0) timeRemaining += HH + " hrs<br>";
      if (MM > 0) timeRemaining += MM + " mins<br>";
      timeRemaining += SS + " secs";

      if (showDiv) document.getElementById("dTimeRemaining").innerHTML = timeRemaining;
//    document.getElementById("pooh").innerHTML = timeRemaining;
      timerId = setTimeout("sessionStatus()",1000);

    }
    
    function sessionRestart() {
      clearTimeout(timerId);
      var vWs = WebService("/V5/SessionControl_ws.asp", "")
      if (vWs == "eof") {
        showAlert('sessionEnd');
      } else {  
        divOff('sessionAlert');
        divOff("divBackground");
        sessionStart();
        sessionStatus();
      }
    }

    //   this displays the appropriate div in the middle of the screen
    function showAlert(theDiv) {
      divOn("divBackground");     
      var divWidth  = 350;
      var divHeight = 200;
      var divTop    = ((document.body.clientHeight - divHeight) / 2) - 50; 
      var divLeft   = ((document.body.clientWidth  - divWidth)  / 2);
      document.getElementById(theDiv).style.width  = divWidth;
      document.getElementById(theDiv).style.height = divHeight;
      document.getElementById(theDiv).style.left   = divLeft;
      document.getElementById(theDiv).style.top    = divTop;
      toggle(theDiv);     
    }
  </script>

</body>

</html>

