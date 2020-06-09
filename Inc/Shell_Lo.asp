          <!--content precedes-->
          <p style="text-align:left">

      </td>
      <td class="shellRight"></td>
    </tr>
    <tr>
      <td class="shellBottomLeft"></td>
      <td class="shellBottom"></td>
      <td class="shellBottomRight"></td>
    </tr>
  </table>



  <% 
    If svMembLevel = 5 Or svHost = "localhost/v5" Then 
    '...these values allow administrators or partner id = 3 (test partner) to see session variables and work the translation engine:
    '           d: display session variables 
    '           t: display session variables and turn ON  translation
    '           x: display session variables and turn OFF translation 
    '           r: refresh this page
    '          en: change language to EN
    '          fr: change language to FR
    '          es: change language to ES
  %>

  <script>
      function ChangeURLTranslate(vTranslate) 
      {
        var vUrl
        if (location.search.length>0)
          if (location.href.indexOf('vTranslate=')>0)
            vUrl=location.href.replace(location.href.substring(location.href.indexOf('vTranslate='),location.href.indexOf('vTranslate=')+12),'vTranslate=' + vTranslate)
          else {
            if (location.href.indexOf('\#')>0)
              vUrl=location.href.substring(0,location.href.indexOf('\#')) + '&vTranslate=' + vTranslate + location.href.substring(location.href.indexOf('\#'))
            else
              vUrl=location.href + '&vTranslate=' + vTranslate
          }
        else
          vUrl=location.href + '?vTranslate=' + vTranslate
        return vUrl
      }
  
      function ChangeURLLang(vNewLang) {
        var vUrl
        if (location.search.length>0)
          if (location.href.indexOf('vLang=')>0)
            vUrl=location.href.replace(location.href.substring(location.href.indexOf('vLang='),location.href.indexOf('vLang=')+8),'vLang=' + vNewLang)
          else
            if (location.href.indexOf('\#')>0)
              vUrl=location.href.substring(0,location.href.indexOf('\#')) + '&vLang=' + vNewLang + location.href.substring(location.href.indexOf('\#'))
            else
              vUrl=location.href + '&vLang=' + vNewLang
        else
          vUrl=location.href + '?vLang=' + vNewLang
        return vUrl
      }  
  </script>

  <%    
    If svTranslate Then 
  %>
  <div style="text-align:center">
    <table>
      <tr>
        <td class="debug" style="text-align:center">
          <a class="debug" href="javascript:location.href=ChangeURLTranslate('n')">Off</a>
          | <a target="_blank" href="/V5/TranslationEngine2.asp?vOk=True&vSelectPages=<%=svPage%>">Translate</a> | <a href="/V5/TranslationEngine1.asp">Engine</a>
          | <a href="#" onclick="window.open('Sessions.asp?vClose=y','Session','toolbar=no,width=450,height=800,left=10,top=10,status=yes,scrollbars=yes,resizable=yes')">Sessions</a>
          | <a class="debug" href="javascript:history.back(1)">Return</a>
          | <a class="debug" href="javascript:location.href=location.href">Refresh</a>
          | <a class="debug" href="javascript:location.href=ChangeURLLang('EN')">EN</a>
          | <a class="debug" href="javascript:location.href=ChangeURLLang('FR')">FR</a>
          | <a class="debug" href="javascript:location.href=ChangeURLLang('ES')">ES</a>
        </td>
      </tr>
    </table>
    <br>
  </div>
  <%
      sHiddenPhrases
    Else
  %>

<p class="debug" style="text-align:center">
  <a class="debug" href="javascript:location.href=ChangeURLTranslate('y')">o</a>&nbsp;  
  	<a class="debug" target="_top" href="/V5/Source/Default.asp?vPage=<%=Session("Page")%>">s</a>&nbsp; 
  	<a class="debug" target="_top" href="/V5/Code/Default.asp?vPage=<%=Session("Page")%>">c</a>&nbsp;&nbsp;&nbsp;&nbsp; 
    <a class="debug" href="javascript:location.href=ChangeURLLang('EN')">EN</a>&nbsp;
    <a class="debug" href="javascript:location.href=ChangeURLLang('FR')">FR</a>&nbsp;
    <a class="debug" href="javascript:location.href=ChangeURLLang('ES')">ES</a>
  <br><%=svCustId%> | <a class="debug" target="_top" href="<%=Session("Page")%>"><%=Session("Page")%></a><br><br></p>

<%
        End If
      End If   

      '... if we are in a secure session then monitor timeout and warn if getting close      
      If Session("Secure") = True Then 
        Session("SessionStarted") = Now()

        '...comment out next line when live
'       Session.TimeOut = 3   '...for testing (must be in whole minutes, start at 6 to show notice at 3, alert at 2 and death at 0) 
%>

<div class="div" id="sessionTracker">
  <div style="text-align: center" class="debug">
    <span id="sessionTime"></span>
  </div>
</div>

<div class="div" id="divBackground" style="height: 100%; width: 100%; position: absolute; left: 1px; top: -1px; background-color: #FFFFFF">
  <div class="div" id="sessionAlert" name="sessionAlert" align="center" style="position: absolute; background-color: #ffffff">
    <table>
      <tr>
        <td class="c5" style="text-align:center">
          <% If svLang = "FR" Then %>
            Votre session expirera <span id="dTimeRemaining">.</span><br><br>Toute donnée non-sauvegardée sera perdu.<br>Si vous avez besoin de plus de temps cliquez ici.<br><br>
          <input onclick="sessionRefresh();" type="button" value="Continuez" name="bContinue" class="button">
          <% Else %>
            Your session will expire in <span id="dTimeRemaining"></span>.<br>Any any unsaved data will be lost.<br><br>If you need more time click here.<br><br>
          <input onclick="sessionRefresh();" type="button" value="Continue" name="bContinue" class="button">
          <% End If %>
      </tr>
    </table>
  </div>

  <div class="div" id="sessionTerminated" name="sessionAlert" align="center" style="position: absolute; background-color: #ffffff">
    <table>
      <tr>
        <td class="c5" style="text-align:center">
          <% If svLang = "FR" Then %>
            Votre session a expiré.<br><br>Toute donnée non-sauvegardée a été perdue.
            <% Else %>
            Your session has expired.<br><br>Any unsaved data has been lost.
            <% End If %>
        </td>
      </tr>
    </table>
  </div>
</div>

<script>   
    var HH, MM, timer, timeRemaining, secsAlert, minsShow, minsAlert, showDiv;

    showDiv = false;
    minsShow  = 10;                      // set to 10 mins for prod
    minsAlert = 5;                       // set to  5 mins for prod
//  minsShow  = <%=Session.Timeout%>;    // set to  2 mins for test
//  minsAlert = 1;                       // set to  5 mins for test
  
  
    // ensure the webservice is available, if so start session, else do not use this feature
    try {  
      sessionCheck();
      sessionStatus();
    } 
    catch (err) {
    }

    // check if how many minutes remaining via the web service
    function sessionCheck() {       
      MM = parseInt(WebService("/V5/SessionRemaining_ws.asp", ""));
    }

    // this is run every minute
    function sessionStatus() {       
      sessionCheck();  // if it works above we assume it will work every time
      //   start counting down (this reduces the session by 1 minute prematurely - which is ok)
      MM = MM - 1;
      //   if end of session then show terminate message
      if (MM == 0) {
        clearTimeout(timer);
        location.href = "#aShell_top";
        divOff('sessionTracker');       
        showAlert('sessionTerminated');     
        return;  
      } 

      timeRemaining = MM + " min";

      //   show timer when alert time is reached (1/2 session time)
      if (MM <= minsShow) {
        divOn('sessionTracker');       
      }  

      //   when we reach this stage alert user
      if (MM <= minsAlert) {
        showAlert('sessionAlert');       
        location.href = "#aShell_top";
        showDiv = true;
      }  

      if (showDiv) {
        document.getElementById("dTimeRemaining").innerHTML = timeRemaining;
      }  

      <% If svLang = "FR" Then %>
        document.getElementById("sessionTime").innerHTML = "La session expire en " + timeRemaining + ".";
      <% Else %>
        document.getElementById("sessionTime").innerHTML = "Session expires in " + timeRemaining + ".";
      <% End If  %>

      // check session status every minute (60,000 milliseconds)
      clearTimeout(timer);
      timer = setTimeout("sessionStatus()",60000);
    }
   

    function sessionRefresh() {
      clearTimeout(timer);
      var vWs = WebService("/V5/SessionRefresh_ws.asp", "")
      if (vWs == "err") {
        showAlert('sessionTerminated');
      } else {  
        divOff('sessionAlert');
        divOff("divBackground");
        sessionCheck();
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
      divOn(theDiv);     
    }
</script>

<% End If %>

<%
    '...Client Versacold cannot handle google analytics
    If Not svSSL And svCustAcctId <> "2814" Then
%>

  <script type="text/javascript">
    // revised Jul 05, 2013
    var _gaq = _gaq || [];
    _gaq.push( ['_setAccount', 'UA-23883721-1'] );
    _gaq.push( ['_trackPageview'] );

    ( function () {
      var ga = document.createElement( 'script' ); ga.type = 'text/javascript'; ga.async = true;
      ga.src = ( 'https:' == document.location.protocol ? 'https://ssl' : '//www' ) + '.google-analytics.com/ga.js';
      var s = document.getElementsByTagName( 'script' )[0]; s.parentNode.insertBefore( ga, s );
    } )();
  </script>

<% 
    End If  
%>