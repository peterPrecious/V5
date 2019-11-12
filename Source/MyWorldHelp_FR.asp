<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body link="#000080" vlink="#000080" alink="#000080" bgcolor="#FFFFFF" text="#000080">

  <% Server.Execute vShellHi %>
  <table border="0" width="90%" id="table1" cellspacing="0" cellpadding="0">
    <tr>
      <td>
      <h1>Bienvenue à &quot;Mes études&quot;</h1>
      <p class="c2">Cette fonction offerte par Vubiz permet à des collègues de travailler en étroite collaboration. La zone d&#39;apprentissage peut être configurée pour refléter la structure de l&#39;organisation cliente. Les apprenants et apprenantes peuvent ainsi être regroupés en fonction de leurs tâches, de leur région ou de leur langue. La zone consacrée à la gestion de projets vous permet de créer des tâches ou des processus qui permettront aux membres de l&#39;équipe de travailler ensemble.<br><br>Les fonctions propres à Mes études sont généralement installées par Vubiz en fonction des besoins des clients. Toutefois, avec une certaine formation, les clients peuvent définir eux-mêmes les tâches à accomplir. Les clients peuvent ainsi assigner les éléments d&#39;actif numériques (par exemple, les modules d&#39;apprentissage) aux tâches appropriées. </p>
      <h1>Services</h1>
      <p class="c2">L&#39;utilisateur ou l&#39;utilisatrice du service Mes études dispose de cinq fonctions principales conviviales. Pour y accéder, il suffit de cliquer sur les icônes qui se retrouvent à droite. L&#39;absence d&#39;icônes indique que les fonctions ne sont pas pertinentes pour l&#39;exercice en question. Voici une brève description des cinq fonctions :</p>
      <div align="center">
        <table border="1" style="border-collapse: collapse" bordercolor="#DDEEF9" width="80%" id="AutoNumber1">
          <tr>
            <td valign="top" width="30"><img border="0" src="../Images/Icons/Email.gif"></td>
            <td valign="top" class="c2" nowrap>
            <h1>Avis par courriel</h1>
            </td>
            <td valign="top" class="c2">Il s&#39;agit d&#39;un outil puissant qui vous permet d&#39;aviser les membres de votre équipe lorsque vous participez aux activités de formation. Vous n&#39;avez qu&#39;à indiquer les personnes que vous désirez aviser et leur laisser un court message expliquant les modifications que vous avez apportées. Ensuite, envoyez le message. Les autres utilisateurs seront avisés d&#39;un ajout sur le site. Il leur suffira de cliquer sur un lien spécial qui les amènera directement dans la section Mes études. Notez bien: ce service devrait être utilisé avec circonspection pour éviter les abus.&nbsp; </td>
          </tr>
          <tr>
            <td valign="top" width="30"><img border="0" src="../Images/Icons/dialogue.gif"></td>
            <td valign="top" class="c2" nowrap>
            <h1>Dialogue</h1>
            </td>
            <td valign="top" class="c2">Vous êtes invité(e) à faire part de vos commentaires et opinions lorsque cet icône apparaît. Inscrivez un bref message. Il sera placé au sommet de la liste de messages déjà envoyés. Voilà! Plusieurs personnes se servent de l&#39;avis par courriel pour aviser les autres utilisateurs du dialogue en cours et les inviter à y participer.</td>
          </tr>
          <tr>
            <td valign="top" width="30"><img border="0" src="../Images/Icons/Notepad.gif"></td>
            <td valign="top" class="c2" nowrap>
            <h1>Notes</h1>
            </td>
            <td valign="top" class="c2">Les notes ressemblent au dialogue. Toutefois, vous êtes la seule personne à les voir. Elles vous permettent de noter vos idées jusqu&#39;à votre retour sur le site.</td>
          </tr>
          <tr>
            <td valign="top" width="30"><img border="0" src="../Images/Icons/Documents.gif"></td>
            <td valign="top" class="c2" nowrap>
            <h1>Documents</h1>
            </td>
            <td valign="top" class="c2">Vous souhaiterez peut-être placer certains documents sur le serveur afin de permettre à vos collègues de les consulter. Le service de documents vous permet de le faire. Les documents sont entreposés et peuvent être téléchargés par les personnes qui désirent les consulter.</td>
          </tr>
          <tr>
            <td valign="top" width="30"><img border="0" src="../Images/Icons/ActionItems.gif"></td>
            <td valign="top" class="c2" nowrap>
            <h1>À faire</h1>
            </td>
            <td valign="top" class="c2">Ce service est particulièrement utile dans le cadre de projets spéciaux ou d&#39;équipes de vente. Il vous permet &quot;d&#39;ordonner&quot; à un ou une collègue d&#39;effectuer une certaine tâche. Les demandes restent ouvertes jusqu&#39;à ce que la personne ayant fait la demande d&#39;action décide de les fermer lorsque la tâche est accomplie.</td>
          </tr>
        </table>
      </div>
      <h1>Le système d&#39;apprentissage en ligne comporte trois différents niveaux d&#39;utilisateurs.</h1>
      <ul class="c2">
        <li>Administrateurs : Les administrateurs et administratrices sont des employés de Vubiz. Ils et elles peuvent apporter des modifications au site pour répondre aux besoins des clients.<br>Les administrateurs et administratrices peuvent modifier certains paramètres de base ou les réactiver au besoin.</li>
        <li>Gestionnaires: Les gestionnaires ont les mêmes fonctions que les animateurs-formateurs et animatrices-formatrices. Ils sont toutefois en mesure de créer et de gérer le site Mes études. <br>Le nombre de gestionnaires sera généralement moins important que le nombre d&#39;animateurs-formateurs ou animatrices-formatrices.</li>
        <li>Animateurs-formateurs: Les animateurs-formateurs et animatrices-formatrices peuvent ajouter de nouvelles inscriptions au site et modifier toutes les informations, sauf l&#39;emplacement.<br>Les animateurs-formateurs et animatrices-formatrices peuvent créer certains rapports d&#39;utilisation en ligne.</li>
        <li>Membres: Les membres peuvent modifier leur nom, leur mot de passe et leur adresse de courriel.&nbsp; Les membres peuvent suivre le cours en ligne et compléter les examens.</li>
      </ul>
      <h1>Pour commencer...</h1>
      <ul class="c2">
        <li>Pour accéder au site d&#39;apprentissage électronique de Vubiz, inscrivez l&#39;adresse www.vubiz.com et sélectionnez la langue de votre choix. Inscrivez votre NIP, qui est composé de quatre lettres et de quatre chiffres (par exemple, ABCD1234) et inscrivez ensuite votre mot de passe. Votre session est maintenant commencée. Vous verrez un mot de bienvenue sur la première page ainsi que des informations sur votre utilisation du site. Vous pouvez modifier les informations en cliquant sur le bouton Édition, qui se retrouve sous &quot;mon profil&quot;. Nous vous suggérons de changer immédiatement votre mot de passe.</li>
        <li>Le contenu du cours se trouvera au bas de la page. Cliquez sur le symbole plus (+) pour ouvrir un dossier. La plupart de ces dossiers comprendront des modules de formation. Pour y accéder, cliquez sur le livre bleu <img border="0" src="../Images/Icons/bookclosed.gif">.</li>
        <li>Lorsque vous avez ouvert un module, vous pouvez vous servir des flèches au bas de l&#39;écran à droite pour naviguer d&#39;une page à l&#39;autre. Le menu déroulant vous permet également de sauter des sections ou de revenir sur vos pas.</li>
        <li>Les modules peuvent comprendre un ou deux tests. Certains modules n&#39;en comprennent pas. La plupart des modules comprennent toutefois une auto-évaluation. Vous verrez un icône dans le coin supérieur droit du module si un test est disponible.<br>&nbsp;</li>
      </ul>
      <div align="center">
        <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber2">
          <tr>
            <td width="100%">
            <ul class="c2">
              <li><img border="0" src="../Images/Icons/Assessment_ON.gif"> Auto-évaluation – cette approche relativement informelle permet à l&#39;utilisateur ou à l&#39;utilisatrice de mettre ses connaissances à l&#39;épreuve.<br>&nbsp;</li>
              <li><img border="0" src="../Images/Icons/Check.gif"> Examen – cette approche est plus formelle et elle est dotée de paramètres stricts qui visent à éviter toute manipulation du système. L&#39;examen comporte la matière du programme en entier. Par exemple, si un programme de certificat comporte huit modules, les questions d&#39;examen porteront sur les huit modules et seront présentées dans le désordre. Vous obtiendrez des précisions supplémentaires sur le site.<br>&nbsp;</li>
              <li>Les résultats des examens seront enregistrés par le système et pourront être consultés par les animatrices-formatrices et animateurs-formateurs. Un certificat attestant que la participante ou le participant a complété le programme sera disponible pour être imprimé à la fin d&#39;un module ou d&#39;un programme de certificat.</li>
            </ul>
            </td>
          </tr>
        </table>
      </div>
      <ul class="c2">
        <li>Les membres et les animateurs-formateurs qui se trouvent dans le même emplacement peuvent s&#39;envoyer des courriels en cliquant sur l&#39;enveloppe <img border="0" src="../Images/Icons/Email.gif">qui se trouve à droite (le cas échéant).</li>
        <li>Lorsque vous avez terminé votre visite sur le site, veuillez mettre fin à votre session en cliquant sur le lien approprié de la barre de menu supérieure.</li>
      </ul>
      <p align="center"><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_FR.gif"></a></p></td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
