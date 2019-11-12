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

<body>

  <% Server.Execute vShellHi %>
  <div align="center">
    <table border="0" width="90%" id="table1" cellspacing="0" cellpadding="0">
      <tr>
        <td>
        <h1>Using &quot;My Learning&quot;</h1>
        <p class="c2">This component of the Vubiz service is where colleagues work together.&nbsp; In the learning space, it allows teams or classrooms to learn in a way that mirrors an organization&#39;s structure, ie by function, region or language.&nbsp; In the project management space, you can setup tasks or processes to help the learners work together in a collaborative manner.&nbsp; Typically &quot;My Learning&quot; is setup by your host personnel based on the clients needs, but clients can, with a little training, assume the responsibility of defining their tasks and assigning the appropriate &quot;digital assets&quot; (ie learning modules, etc) to these tasks. </p>
        <h1>Services</h1>
        <p class="c2">To make this world more dynamic five &quot;services&quot; are available to the learner in &quot;My Learning&quot;.&nbsp; These services are accessible by specific icons that may appear on the right side of the task list.&nbsp; If none appear, that means that they were not deemed relevant to the task at hand.&nbsp; Here&#39;s what they can offer you:<br><br>&nbsp;</p>
        <div align="center">
          <table border="1" style="border-collapse: collapse" bordercolor="#DDEEF9" width="80%" id="AutoNumber1" cellspacing="0" cellpadding="3">
            <tr>
              <td valign="top" width="30"><img border="0" src="../Images/Icons/Email.gif"></td>
              <th width="80" valign="top" nowrap align="left" class="c1">Email Alert</th>
              <td valign="top" class="c2">This is a powerful tool.&nbsp; Every member of the team is available to &quot;alert&quot; when you make a contribution to the site.&nbsp; You simply check off who you want to contact, leave a simple message informing the learner(s) of the additions you&#39;ve make, and click send.&nbsp; What the learner receives is an alert that there are items of interest at the site.&nbsp; The learner will then simply click on a special URL and be brought right into &quot;My Learning&quot;.&nbsp; Note: Do not over use this service - excessive emailing is annoying.&nbsp; </td>
            </tr>
            <tr>
              <td valign="top" width="30"><img border="0" src="../Images/Icons/dialogue.gif"></td>
              <th width="80" valign="top" nowrap align="left" class="c1">Dialogue</th>
              <td valign="top" class="c2">Where ever you see this icon, you are encouraged to share your thoughts.&nbsp; Simply type a brief message and it will go to the top of the dialogue &quot;thread&quot;.&nbsp; That&#39;s it!&nbsp; (Oh, this is why you tend to use the email alert, so others can read your dialogue.</td>
            </tr>
            <tr>
              <td valign="top" width="30"><img border="0" src="../Images/Icons/Notepad.gif"></td>
              <th width="80" valign="top" nowrap align="left" class="c1">Notes</th>
              <td valign="top" class="c2">Notes are similar to Dialogue but no one sees them except yourself.&nbsp; They are just way for your to record your thoughts on the assumption that you will return later to continue.</td>
            </tr>
            <tr>
              <td valign="top" width="30"><img border="0" src="../Images/Icons/Documents.gif"></td>
              <th width="80" valign="top" nowrap align="left" class="c1">Documents</th>
              <td valign="top" class="c2">Sometimes you may wish to &quot;upload&quot; a document to the server so other colleagues can access them.&nbsp; This service provides that functionality.&nbsp; Documents sit in a &quot;Repository&quot; available for anyone to &quot;download&quot; to their computer to study.</td>
            </tr>
            <tr>
              <td valign="top" width="30"><img border="0" src="../Images/Icons/ActionItems.gif"></td>
              <th width="80" valign="top" nowrap align="left" class="c1">Action Items</th>
              <td valign="top" class="c2">This is a powerful service particularly suited for projects or sales teams.&nbsp; It allows you to &quot;order&quot; a colleague to perform a task.&nbsp; This is an action item that remains &quot;open&quot; until completed by the designated &quot;owner&quot; of the task.</td>
            </tr>
          </table>
        </div>
        <h1>The E-Learning Platform has different levels with respect to learners.</h1>
        <ul class="c2">
          <li>Administrators : who are your &quot;hosts&quot; who setup this site based on your organizations requirements. </li>
          <li>Managers: who have special reporting rights and can setup &quot;facilitators&quot;.&nbsp; There should one be one &quot;manager&quot; per site. </li>
          <li>Facilitators: who add new members to the site and have the ability to generate the basic on-line reports. </li>
          <li>Members: who can modify their name and email address when they enter the site and can study on-line and take the corresponding tests. </li>
        </ul>
        <h1>How to Begin</h1>
        <ul class="c2">
          <li>Once you enter the site, you will see on the first page a welcome statement with information on your personal usage. For example: Welcome Sally. If you click on the edit button under &quot;my profile&quot; you may change your information. </li>
          <li>The learning content can be found either under &quot;My Content&quot; or in &quot;My Learning&quot;.&nbsp; Click on the plus sign to open an area. Most of these areas will contain learning modules. To access these, click on the blue book <img border="0" src="../Images/Icons/bookclosed.gif">. </li>
          <li>Once inside a module use the arrows at the bottom right of the screen to navigate page by page or choose a section from the drop down menu to skip ahead or review a past section. </li>
          <li>Some modules will have one of two testing applications or none at all. Most modules will contain a self assessment. When you are on the last page of a module an icon will appear at the top right corner of the module.</li>
        </ul>
        <div align="center">
          <table border="0" cellpadding="0" style="border-collapse: collapse" bordercolor="#111111" width="80%" id="AutoNumber2">
            <tr>
              <td width="100%">
              <ul class="c2">
                <li><img border="0" src="../Images/Icons/Assessment_ON.gif"> <b>Self Assessment</b>:&nbsp; this is a less constrained method for the learner to test their comprehension of the module.<br>&nbsp;</li>
                <li><img border="0" src="../Images/Icons/Check.gif"> <b>Examination</b>: this is a more robust method with strict parameters intended to prevent manipulation of the system. The examination covers the content of the entire program. For instance: If a certificate program has 8 modules, the exam questions will reflect the content from all 8 modules in no particular order. Further details are found within the site.<br>&nbsp;</li>
                <li>All test results will be tracked by the system and available for viewing by Facilitators and Managers. A printable certificate of completion will be generated upon successful completion of a module / certificate program.</li>
              </ul>
              </td>
            </tr>
          </table>
        </div>
        <ul class="c2">
          <li>Members and facilitators in a specified location may email alert each other using the envelope icon <img border="0" src="../Images/Icons/Email.gif"> to the right of any area (if available). </li>
          <li>When you are done using the site, please use the sign-off link on the top menu bar to disconnect from the system. </li>
        </ul>
        <p align="center"><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a> </p></td>
      </tr>
    </table>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


