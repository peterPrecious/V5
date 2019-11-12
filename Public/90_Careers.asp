<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<html>

<head>
  <title>Welcom</title>
  <link href="Resources/Public.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>
    var msg = "Please remember to include your name at the bottom of the survey."
    function jSurvey(modid) {
      var url = "http://vubiz.com/V5/Default.asp?vLang=EN&vCust=DEMO1001&vId=CAREER&vQModId="+modid;
      modwindow = window.open(url,'Module','toolbar=no,location=1,width=770,height=500,left=50,top=50,status=yes,scrollbars=yes,resizable=yes');
    }
  </script>
  <base target="_self">
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <table border="0" width="100%" cellspacing="0" cellpadding="10" id="table3">
    <tr>
      <td class="c2" valign="top">
      <h1 align="center">Enjoy an IT Career at&nbsp; VUBIZ!</h1>
      <blockquote>
        <table border="0" cellpadding="0" cellspacing="0" bordercolor="#DDEEF9" id="table6" width="80%">
          <tr>
            <td valign="top">
            <p class="c2">Vubiz is one of Canada?s leading e-learning firms providing web based training and consulting services to a wide range of government, corporate and association clients. We are looking for a senior web systems developer to join our tight, innovative, dedicated team based in Oakville, Ontario.</p>
            <p class="c2">You will be expected to analyze, design, code, test and implement your web applications. You should have at least 3 to 5 years experience working with Classic ASP, ASP.Net, C# (preferred), SQL Server and JavaScript. Client-facing experience would be an asset.</p>
            <p class="c2">Note: this is not a managerial position - you must have excellent system analysis skills and be a &quot;hard core coder&quot;.</p>
            <p class="c2">If you are motivated, innovative, and enjoy the challenge of working in a fast-growing environment, you are a match for Vubiz.&nbsp; Growth opportunities in this position are very strong.</p>
            <p class="c2">To apply:</p>
            <ol class="c2">
              <li>Please complete a brief online survey by <a class="c2" href="javascript:alert(msg);jSurvey('9972EN')">clicking here</a> (<font color="#FF0000">do NOT forget to put your name in the text box at the end of the survey</font>);<br>&nbsp;</li>
              <li>Please email your MS Word or PDF resume to: <a href="mailto:jobs@vubiz.com?subject=IT Career Application">jobs@vubiz.com</a>.</li>
            </ol>
            <p class="c2">Peter Bulloch will respond via email as soon as possible via <a href="mailto:pbulloch@vubiz.com?subject=Job Enquiry">pbulloch@vubiz.com</a>.&nbsp; Ensure this email address can get through any filters.</p>
            <p class="c2">Thank you for considering Vubiz.&nbsp;</p>
            <p class="c2">&nbsp;</p>
            </td>
          </tr>
          </td>
          </tr>
        </table>
      </blockquote>
      </td>
    </tr>
  </table>

</body>

</html>
