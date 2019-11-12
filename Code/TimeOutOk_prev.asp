<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% Session("Secure") = False %>


<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Vubiz Inactive Session</title>
</head>

<body>

  <% Server.Execute vShellHi %>
  <div align="center">
    <table cellpadding="0" border="0" style="border-collapse: collapse" bordercolor="#111111" width="80%">
      <tr>
        <th width="100%" align="left">
        <h1 align="center"><!--webbot bot='PurpleText' PREVIEW='You have lost connection with the Vubiz service'--><%=fPhra(000045)%></h1>
        <h2><!--webbot bot='PurpleText' PREVIEW='Whenever the Vubiz service cannot communicate properly with your browser it will disconnect. Here are some of the reasons:'--><%=fPhra(000032)%></h2>
        <ul class="c2">
          <li><!--webbot bot='PurpleText' PREVIEW='You left the service unattended for 20 minutes or more. This will cause a Timeout disconnection.'--><%=fPhra(000049)%></li>
          <li><!--webbot bot='PurpleText' PREVIEW='You temporarily lost your Internet connection.'--><%=fPhra(000686)%></li>
          <li><!--webbot bot='PurpleText' PREVIEW='You have set your browser security settings too tight. The default settings are normally perfect for running the Vubiz service. If they have been &quot;tightened&quot; you may not be able to communicate properly with Vubiz. These browser settings are in &quot;Tools/Internet Options/Security&quot; (may vary with different versions).'--><%=fPhra(000047)%></li>
          <li><!--webbot bot='PurpleText' PREVIEW='You clicked the &quot;Back button&quot; on your browser at the wrong spot. It is best NEVER to use your browser back button.'--><%=fPhra(000012)%></li>
          <li><!--webbot bot='PurpleText' PREVIEW='You have use of &quot;Cookies&quot; disabled on your browser. Ensure you set your browser to enable cookies. The browser settings are in &quot;Tools/Internet Options/Security&quot; (may vary with different versions).'--><%=fPhra(000048)%></li>
          <li><!--webbot bot='PurpleText' PREVIEW='You accessed Vubiz within a &quot;pseudo&quot; browser (for example, within Outlook). To access Vubiz you must open the Internet Explorer browser and access our service from there.'--><%=fPhra(000025)%></li>
          <li><!--webbot bot='PurpleText' PREVIEW='You right-clicked on a link within the site which will open that page in a new window. Never right click on a link as the service needs all the component frames to work properly.'--><%=fPhra(000050)%></li>
          <li><!--webbot bot='PurpleText' PREVIEW='You launched the site from an invalid bookmark. You must start the service at the beginning where you signed in - not at a page within the service.'--><%=fPhra(000016)%></li>
        </ul>
        <h2><!--webbot bot='PurpleText' PREVIEW='Please sign-into this service again to resume your learning.'--><%=fPhra(000200)%> <br>&nbsp; </h2></th>
      </tr>
    </table>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

<!--
UPDATE Phra SET Phra_EN = '<h1 align="center">You have lost connection with the Vubiz service</h1><p class="c2">Whenever the Vubiz service cannot communicate properly with your browser it will disconnect. Here are some of the reasons:<ul><li><p class="c2">You left the service unattended for 20 minutes or more. This will cause a Timeout disconnection.</li><li><p class="c2">You have set your browser security settings too tight. The default settings are normally perfect for running the Vubiz service. If they have been "tightened" you may not be able to communicate properly with Vubiz. These browser settings are in &quot;Tools/Internet Options/Security&quot; (may vary with different versions).</li><li><p class="c2">You clicked the &quot;Back button&quot; on your browser at the wrong spot. It is best NEVER to use your browser back button.</li><li><p class="c2">You have use of &quot;Cookies&quot; disabled on your browser. Ensure you set your browser to enable cookies. The browser settings are in &quot;Tools/Internet Options/Security&quot; (may vary with different versions).</li><li><p class="c2">You accessed Vubiz within a &quot;pseudo&quot; browser (for example, within Outlook). To access Vubiz you must open the Internet Explorer browser and access our service from there.</li><li><p class="c2">You right-clicked on a link within the site which will open that page in a new window. Never right click on a link as the service needs all the component frames to work properly.</li><li><p class="c2">You launched the site from an invalid bookmark. You must start the service at the beginning where you signed in - not at a page within the service.</li></ul><p class="c2">Please sign-into this service again to resume your learning.' WHERE (Phra_No = 1031)
UPDATE Phra SET Phra_FR = '<h1 align="center">Votre connexion avec le service Vubiz a été interrompue.</h1><p class="c2">Le service Vubiz se débranche lorsqu’il ne peut communiquer correctement avec votre navigateur. Voici certaines des raisons possibles :<ul><li><p class="c2">Vous n’avez pas utilisé le service durant 20 minutes ou plus. Cela cause un débranchement par délai d’inactivité.</li><li><p class="c2">Les mesures de sécurité de votre navigateur sont trop restrictives. Le service Vubiz fonctionne habituellement de façon optimale avec les paramètres par défaut. Si les mesures de sécurité sont plus restrictives, vous pourriez avoir de la difficulté à communiquer avec Vubiz. Vous trouverez les paramètres dans «Tools/Internet Options/Security» (cela peut varier d’un logiciel à l’autre).</li><li><p class="c2">Vous avez cliqué sur le bouton Page précédente de votre navigateur au mauvais endroit. Il est préférable de ne JAMAIS utiliser le bouton Page précédente.</li><li><p class="c2">Les témoins sont désactivés dans votre navigateur. Assurez-vous que les témoins sont activés. Vous trouverez les paramètres dans «Tools/Internet Options/Security» (cela peut varier d’un logiciel à l’autre).</li><li><p class="c2">Vous avez accédé à Vubiz à l’aide d’un «pseudo-navigateur» (par exemple, à même le logiciel Outlook). Pour accéder au service Vubiz, vous devez utiliser le navigateur Internet Explorer.</li><li><p class="c2">Vous avez cliqué à droite sur un lien dans le site qui ouvre une page dans une nouvelle fenêtre. Ne cliquez jamais à droite sur un lien car tous les cadres sont nécessaires pour que le service fonctionne correctement.</li><li><p class="c2">Vous avez accédé au site à partir d’un signet invalide. Vous devez accéder au service en ouvrant votre session à la page d’accueil - et non à une autre page à l’intérieur du système.</li></ul><p class="c2"><br>Veuillez ouvrir une autre session pour continuer votre apprentissage.' WHERE (Phra_No = 1031)
-->



