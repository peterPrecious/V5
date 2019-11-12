<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<html><head><title></title><meta name="GENERATOR" content="Microsoft FrontPage 5.0"><meta name="ProgId" content="FrontPage.Editor.Document"></head>
<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

<% Server.Execute vShellHi %>


 <table width="100%" border="1" bordercolor="#C6DBFF" style="border-collapse: collapse" cellspacing="0" cellpadding="3">
   <tr>
     <td colspan="2" align="center"><b><font face="Verdana" size="1">Vubiz WebService Functionality<br><br>&nbsp;</font></b></td>
   </tr>
   <tr>
     <td align="right" width="150" valign="top"><b><font face="Verdana" size="1">Request Type :&nbsp; </font></b></td>
     <td valign="top"><font face="Verdana" size="1">GetCatalogue</font></td>
   </tr>
   <tr>
     <td align="right" width="150" valign="top"><b><font face="Verdana" size="1">Returns :&nbsp; </font></b></td>
     <td valign="top"><font face="Verdana" size="1">Full Courseware Catalogue as setup for this account.</font></td>
   </tr>
   <tr>
     <td align="right" width="150" valign="top"><b><font face="Verdana" size="1">Example :&nbsp; </font></b></td>
     <td valign="top"><font face="Verdana" size="1">&lt;VUBIZ&gt;&lt;WS vAction=&#39;GetCatalogue&#39; vCust=&#39;ABCD1234&#39; vId=&#39;ADMINISTRATOR&#39;/&gt;&lt;/VUBIZ&gt;</font></td>
   </tr>
   <tr>
     <td align="right" width="150" valign="top"><b><font face="Verdana" size="1">Requires :&nbsp; </font></b></td>
     <td valign="top">
     <ul>
       <li><font face="Verdana" size="1">valid Customer (vCust) </font></li>
       <li><font face="Verdana" size="1">valid Password Id (vId)</font></li>
     </ul>
     </td>
   </tr>
   <tr>
     <td align="right" width="150" valign="top">&nbsp;</td>
     <td valign="top">&nbsp;</td>
   </tr>
   <tr>
     <td align="right" width="150" valign="top"><font face="Verdana" size="1"><b>Request Type :&nbsp; </b></font></td>
     <td valign="top"><font face="Verdana" size="1">EnrollUser</font></td>
   </tr>
   <tr>
     <td align="right" width="150" valign="top"><b><font face="Verdana" size="1">Returns :&nbsp; </font></b></td>
     <td valign="top"><font face="Verdana" size="1">Enrolls a new user into this account and returns the users password (unless one has been submitted)</font></td>
   </tr>
   <tr>
     <td align="right" width="150" valign="top"><font face="Verdana" size="1"><b>Example :&nbsp; </b></font></td>
     <td valign="top"><font face="Verdana" size="1">&lt;VUBIZ&gt;&lt;WS vAction=&#39;Enroll&#39; vCust=&#39;ABCD1234&#39; vId=&#39;ADMINISTRATOR&#39; vPrograms='P1234EN P1245EN' vExpires='06/15/02' /&gt;&lt;/VUBIZ&gt;</font></td>
   </tr>
   <tr>
     <td align="right" width="150" valign="top"><b><font face="Verdana" size="1">Requires :&nbsp; </font></b></td>
     <td valign="top">
     <ul>
       <li><font face="Verdana" size="1">mandatory, valid Customer (vCust). </font></li>
       <li><font face="Verdana" size="1">optional Password Id (vId) - if the Id is present this user will be enrolled into the site under that Id, if not, a unique password Id will be generated and returned.&nbsp; </font></li>
       <li><font face="Verdana" size="1">optional, valid Programs (vPrograms) - if present, the program(s), separated by spaces, will be added to any other programs that may be available to this user, if not present, this user will get whatever programs were assigned to this account.&nbsp; </font></li>
       <li><font face="Verdana" size="1">optional expiry date for the programs (vExpires as MM/DD/YY), if present this means that the user can access the programs until that date - typically 90 or 365 days from date of order/setup.</font></li>
     </ul>
     </td>
   </tr>
   <tr>
     <td colspan="2"><font face="Verdana" size="1"><br><br>Note.&nbsp; Web Service requests are tracked.&nbsp; Should repeated unsuccessful attempts occur, the service will automatically be disabled.</font></td>
   </tr>
 </table>

<% Server.Execute vShellLo %>


</body></html>