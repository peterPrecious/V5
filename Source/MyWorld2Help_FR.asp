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
      <h1>Bienvenue � &quot;Mes �tudes&quot;</h1>
      <p class="c2">Cette fonction offerte par Vubiz permet � des coll�gues de travailler en �troite collaboration. La zone d&#39;apprentissage peut �tre configur�e pour refl�ter la structure de l&#39;organisation cliente. Les apprenants et apprenantes peuvent ainsi �tre regroup�s en fonction de leurs t�ches, de leur r�gion ou de leur langue. La zone consacr�e � la gestion de projets vous permet de cr�er des t�ches ou des processus qui permettront aux membres de l&#39;�quipe de travailler ensemble.<br><br>Les fonctions propres � Mes �tudes sont g�n�ralement install�es par Vubiz en fonction des besoins des clients. Toutefois, avec une certaine formation, les clients peuvent d�finir eux-m�mes les t�ches � accomplir. Les clients peuvent ainsi assigner les �l�ments d&#39;actif num�riques (par exemple, les modules d&#39;apprentissage) aux t�ches appropri�es. </p>
      <h1>Services</h1>
      <p class="c2">L&#39;utilisateur ou l&#39;utilisatrice du service Mes �tudes dispose de cinq fonctions principales conviviales. Pour y acc�der, il suffit de cliquer sur les ic�nes qui se retrouvent � droite. L&#39;absence d&#39;ic�nes indique que les fonctions ne sont pas pertinentes pour l&#39;exercice en question. Voici une br�ve description des cinq fonctions :</p>
      <div align="center">
        <table border="1" style="border-collapse: collapse" bordercolor="#DDEEF9" width="80%" id="AutoNumber1">
          <tr>
            <td valign="top" width="30"><img border="0" src="../Images/Icons/Email.gif"></td>
            <td valign="top" class="c2" nowrap>
            <h1>Avis par courriel</h1>
            </td>
            <td valign="top" class="c2">Il s&#39;agit d&#39;un outil puissant qui vous permet d&#39;aviser les membres de votre �quipe lorsque vous participez aux activit�s de formation. Vous n&#39;avez qu&#39;� indiquer les personnes que vous d�sirez aviser et leur laisser un court message expliquant les modifications que vous avez apport�es. Ensuite, envoyez le message. Les autres utilisateurs seront avis�s d&#39;un ajout sur le site. Il leur suffira de cliquer sur un lien sp�cial qui les am�nera directement dans la section Mes �tudes. Notez bien: ce service devrait �tre utilis� avec circonspection pour �viter les abus.&nbsp; </td>
          </tr>
          <tr>
            <td valign="top" width="30"><img border="0" src="../Images/Icons/dialogue.gif"></td>
            <td valign="top" class="c2" nowrap>
            <h1>Dialogue</h1>
            </td>
            <td valign="top" class="c2">Vous �tes invit�(e) � faire part de vos commentaires et opinions lorsque cet ic�ne appara�t. Inscrivez un bref message. Il sera plac� au sommet de la liste de messages d�j� envoy�s. Voil�! Plusieurs personnes se servent de l&#39;avis par courriel pour aviser les autres utilisateurs du dialogue en cours et les inviter � y participer.</td>
          </tr>
          <tr>
            <td valign="top" width="30"><img border="0" src="../Images/Icons/Notepad.gif"></td>
            <td valign="top" class="c2" nowrap>
            <h1>Notes</h1>
            </td>
            <td valign="top" class="c2">Les notes ressemblent au dialogue. Toutefois, vous �tes la seule personne � les voir. Elles vous permettent de noter vos id�es jusqu&#39;� votre retour sur le site.</td>
          </tr>
          <tr>
            <td valign="top" width="30"><img border="0" src="../Images/Icons/Documents.gif"></td>
            <td valign="top" class="c2" nowrap>
            <h1>Documents</h1>
            </td>
            <td valign="top" class="c2">Vous souhaiterez peut-�tre placer certains documents sur le serveur afin de permettre � vos coll�gues de les consulter. Le service de documents vous permet de le faire. Les documents sont entrepos�s et peuvent �tre t�l�charg�s par les personnes qui d�sirent les consulter.</td>
          </tr>
          <tr>
            <td valign="top" width="30"><img border="0" src="../Images/Icons/ActionItems.gif"></td>
            <td valign="top" class="c2" nowrap>
            <h1>� faire</h1>
            </td>
            <td valign="top" class="c2">Ce service est particuli�rement utile dans le cadre de projets sp�ciaux ou d&#39;�quipes de vente. Il vous permet &quot;d&#39;ordonner&quot; � un ou une coll�gue d&#39;effectuer une certaine t�che. Les demandes restent ouvertes jusqu&#39;� ce que la personne ayant fait la demande d&#39;action d�cide de les fermer lorsque la t�che est accomplie.</td>
          </tr>
        </table>
      </div>
      <h1>Le syst�me d&#39;apprentissage en ligne comporte trois diff�rents niveaux d&#39;utilisateurs.</h1>
      <ul class="c2">
        <li>Administrateurs : Les administrateurs et administratrices sont des employ�s de Vubiz. Ils et elles peuvent apporter des modifications au site pour r�pondre aux besoins des clients.<br>Les administrateurs et administratrices peuvent modifier certains param�tres de base ou les r�activer au besoin.</li>
        <li>Gestionnaires: Les gestionnaires ont les m�mes fonctions que les animateurs-formateurs et animatrices-formatrices. Ils sont toutefois en mesure de cr�er et de g�rer le site Mes �tudes. <br>Le nombre de gestionnaires sera g�n�ralement moins important que le nombre d&#39;animateurs-formateurs ou animatrices-formatrices.</li>
        <li>Animateurs-formateurs: Les animateurs-formateurs et animatrices-formatrices peuvent ajouter de nouvelles inscriptions au site et modifier toutes les informations, sauf l&#39;emplacement.<br>Les animateurs-formateurs et animatrices-formatrices peuvent cr�er certains rapports d&#39;utilisation en ligne.</li>
        <li>Membres: Les membres peuvent modifier leur nom, leur mot de passe et leur adresse de courriel.&nbsp; Les membres peuvent suivre le cours en ligne et compl�ter les examens.</li>
      </ul>
      <h1>Pour commencer...</h1>
      <ul class="c2">
        <li>Pour acc�der au site d&#39;apprentissage �lectronique de Vubiz, inscrivez l&#39;adresse www.vubiz.com et s�lectionnez la langue de votre choix. Inscrivez votre NIP, qui est compos� de quatre lettres et de quatre chiffres (par exemple, ABCD1234) et inscrivez ensuite votre mot de passe. Votre session est maintenant commenc�e. Vous verrez un mot de bienvenue sur la premi�re page ainsi que des informations sur votre utilisation du site. Vous pouvez modifier les informations en cliquant sur le bouton �dition, qui se retrouve sous &quot;mon profil&quot;. Nous vous sugg�rons de changer imm�diatement votre mot de passe.</li>
        <li>Le contenu du cours se trouvera au bas de la page. Cliquez sur le symbole plus (+) pour ouvrir un dossier. La plupart de ces dossiers comprendront des modules de formation. Pour y acc�der, cliquez sur le livre bleu <img border="0" src="../Images/Icons/bookclosed.gif">.</li>
        <li>Lorsque vous avez ouvert un module, vous pouvez vous servir des fl�ches au bas de l&#39;�cran � droite pour naviguer d&#39;une page � l&#39;autre. Le menu d�roulant vous permet �galement de sauter des sections ou de revenir sur vos pas.</li>
        <li>Les modules peuvent comprendre un ou deux tests. Certains modules n&#39;en comprennent pas. La plupart des modules comprennent toutefois une auto-�valuation. Vous verrez un ic�ne dans le coin sup�rieur droit du module si un test est disponible.<br>&nbsp;</li>
      </ul>
      <div align="center">
        <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber2">
          <tr>
            <td width="100%">
            <ul class="c2">
              <li><img border="0" src="../Images/Icons/Assessment_ON.gif"> Auto-�valuation � cette approche relativement informelle permet � l&#39;utilisateur ou � l&#39;utilisatrice de mettre ses connaissances � l&#39;�preuve.<br>&nbsp;</li>
              <li><img border="0" src="../Images/Icons/Check.gif"> Examen � cette approche est plus formelle et elle est dot�e de param�tres stricts qui visent � �viter toute manipulation du syst�me. L&#39;examen comporte la mati�re du programme en entier. Par exemple, si un programme de certificat comporte huit modules, les questions d&#39;examen porteront sur les huit modules et seront pr�sent�es dans le d�sordre. Vous obtiendrez des pr�cisions suppl�mentaires sur le site.<br>&nbsp;</li>
              <li>Les r�sultats des examens seront enregistr�s par le syst�me et pourront �tre consult�s par les animatrices-formatrices et animateurs-formateurs. Un certificat attestant que la participante ou le participant a compl�t� le programme sera disponible pour �tre imprim� � la fin d&#39;un module ou d&#39;un programme de certificat.</li>
            </ul>
            </td>
          </tr>
        </table>
      </div>
      <ul class="c2">
        <li>Les membres et les animateurs-formateurs qui se trouvent dans le m�me emplacement peuvent s&#39;envoyer des courriels en cliquant sur l&#39;enveloppe <img border="0" src="../Images/Icons/Email.gif">qui se trouve � droite (le cas �ch�ant).</li>
        <li>Lorsque vous avez termin� votre visite sur le site, veuillez mettre fin � votre session en cliquant sur le lien appropri� de la barre de menu sup�rieure.</li>
      </ul>
      <p align="center"><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_FR.gif"></a></p></td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
