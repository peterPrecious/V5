<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/QueryString.asp"-->
<!--#include virtual = "V5/Repository/Documents/EcomDocumentRoutines.asp"-->

<%
  Dim vNoSource, vDocUrl
  If Len(svCustReturnUrl) = 0 Then vNoSource = True Else vNoSource = False
  vDocUrl = fGetDocument ("MultiUserManual")
%>

<html>

<head>
  <title>Ecom3DisplayIds</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <div style="text-align:center">
    <table style="width:600px; margin:auto;">

      <tr>
        <td style="text-align:center">
          <h1><!--[[-->Thank you<!--]]-->.</h1>
          <% If Len(Session("EcomId")) = 0 And Session("Ecom_Media") = "AddOn2" And svMembLevel > 3 Then %>
          <h2>Your purchase was successful.</h2>
          <p><input type="button" onclick="location.href = 'Default.asp'" value="Return" name="bReturn" class="button"></p>
          <% Else %>
          <h2><!--[[-->Here is your new Customer Id and Password <br>which you will need to access your new service.<!--]]--></h2>

          <table style="width:300px; margin:auto;">
            <tr>
              <th style="width:50%"><!--[[-->Customer Id<!--]]--> : </th>
              <td style="width:50%"><%=Session("EcomCust")%></td>
            </tr>
            <tr>
              <th><!--[[-->Password<!--]]--> : </th>
              <td><%=Session("EcomId")%></td>
            </tr>
            <% If vDocUrl <> "" Then %>
            <tr>
              <th style="text-align:center;" colspan="2"><br><a <%=fstatx%> target="_blank" href="<%=vDocUrl%>"><!--[[-->Click to download Facilitator Manual<!--]]--></a></th>
            </tr>
            <% End If%>
          </table>

          <br />

          <h6><!--[[-->PLEASE CLICK THE ABOVE LINK TO DOWNLOAD AND SAVE THE FACILITATOR MANUAL TO YOUR COMPUTER.&ensp;IT CONTAINS INSTRUCTIONS ON OPERATING THIS SERVICE.<!--]]--></h6>
          <br />
          <h3 style="text-align:left;"><!--[[-->Remember that you are the Facilitator for this service and the above Password gives you advanced rights to setup your learners and monitor their progress.&nbsp; When a learner accesses this site, they will see the program(s) you have assigned to them.&nbsp; When you enter, you will see an Administration tab containing links to add/edit learners, monitor their usage of programs and review their performance on assessments.&nbsp; Good luck!<!--]]--></h3>
         
          <h2><br /><!--[[-->To begin click <b>Continue</b> where you can enter your new Customer Id and Password.<!--]]--></h2>
          <p><input type="button" onclick="location.href='<%=fIf(vNoSource, "//" & svHost, svCustReturnUrl) %>'" value="<%="<!--{{-->Continue<!--}}-->"%>" name="B2" class="button"></p>

          <h2><br /><!--[[-->You can also click below to automatically <b>Sign In</b>.<!--]]--></h2>

          <p style="text-align:center"><input onclick="location.href='//<%=svHost%>/default.asp?vCust=<%=Session("EcomCust")%>&vId=<%=Session("EcomId")%>'" type="button" value="<%="<!--{{-->Sign In<!--}}-->"%>" name="B1" class="button"></p>

          <br /><br />
          <table class="table" style="width:300px; margin:auto">
            <tr>
              <td class="rowshade"><a <%=fstatx%> href="javascript:window.print();"><img border="0" src="../Images/Icons/Printer.gif"></a></td>
              <td class="rowshade" style="text-align:left;"><!--[[-->Print this page for your records.<!--]]--></td>
            </tr>
            <tr>
              <td class="rowshade"><img border="0" src="../Images/Icons/Bang.gif"></td>
              <td class="rowshade" style="text-align:left;"><!--[[-->Remember to <b>Sign Off</b> after every session.<!--]]--></td>
            </tr>
            <tr>
              <td class="rowshade"><a <%=fstatx%> href="mailto:<%= fIf(Len(svCustEmail) > 0, svCustEmail, "support@vubiz.com")%>?subject=Ecommerce Issue"><img border="0" src="../Images/Icons/Email3.gif"></a></td>
              <td class="rowshade" style="text-align:left;"><!--[[-->Feel free to email us if you have any questions.<!--]]--></td>
            </tr>
          </table>

        <% End If  %>

        </td>
      </tr>
    </table>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
