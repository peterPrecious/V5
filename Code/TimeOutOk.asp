<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% Session("Secure") = False %>

<html>

  <head>
    <meta charset="UTF-8">
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
    <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
    <title>Vubiz Inactive Session</title>
    <style>.section { border:1px solid black; padding:10px;}</style>
  </head>

  <body>

    <% Server.Execute vShellHi %>
    <table class="table">

      <tr>
        <td class="section">
          <h1 align="center">You have lost connection with the Vubiz service</h1>
          <p>Here are some of the reasons:</p>
          <ul>
            <li>You left the service unattended for 20 minutes or more.</li>
            <li>You temporarily lost your Internet connection.</li>
            <li>You were accessing this service in two separate windows and Signed Off in one of them.</li>
          </ul>
          <p>Please sign-into this service again to resume your learning.<br /><br /></p>
        </td>
      </tr>

      <tr>
        <td class="section">
          <h1 align="center">Vous avez perdu la connexion avec le service Vubiz </h1>
          <p>Voici quelques-unes des raisons : </p>
          <ul>
            <li>vous avez laissé le service sans assistance pendant 20 minutes ou plus. </li>
            <li>vous avez temporairement perdu votre connexion Internet. </li>
            <li>vous ont accès à ce service dans deux fenêtres distinctes et signé Off dans l'un d'entre eux. </li>
          </ul>
          <p>S'il vous plaît connecter à ce service à nouveau pour reprendre votre apprentissage.<br /><br /></p>
        </td>
      </tr>

      <tr>
        <td class="section">
          <h1 align="center">Han perdido conexión con el servicio de Vubiz </h1>
          <p>Aquí es algunas de las razones: </p>
          <ul>
            <li>se dejó el servicio desatendido durante 20 minutos o más. </li>
            <li>has perdido temporalmente su conexión de Internet. </li>
            <li>que fueron accediendo a este servicio en dos ventanas separadas y firmado fuera en uno de ellos. </li>
          </ul>
          <p>Por favor signo-en este servicio otra vez para reanudar su de aprendizaje.<br /><br /></p>
        </td>
      </tr>


    </table>


    <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  </body>

</html>



