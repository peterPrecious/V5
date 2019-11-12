<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <style>.Div {DISPLAY: none; MARGIN: 0px}</style>
  <script src="/V5/Inc/Launch.js"></script>
</head>

<body>

  <% 
  	Server.Execute vShellHi
  %>
  <table cellspacing="5" cellpadding="10" border="0" id="table176" width="100%">
    <form method="POST" action="RC_Email.asp">
      <tr valign="top">
        <td width="100%"><h1 align="center">Welcome to the Resource Centre.</h1><p class="c2">This provides you with a powerful on-line learning tool that supports your delivery of advice and information to your colleagues.&nbsp; This resource centre allows you to preview the modules and deliver the ones you believe will be most useful, directly via email.&nbsp; Upon receipt of the email, your colleague will be directed to this space via an email link where they can review the content you have selected.&nbsp; Click on the Program Title to display the associated modules and click on the Module Title to view its content.</p></td>
      </tr>
      <tr valign="top">
        <td width="100">
        <table cellspacing="1" cellpadding="2">
          <% 
            sGetCust (svCustId)
            Dim aProg, aMods
            aProg = Split(vCust_Resources, " ")
            For i = 0 to Ubound(aProg)
              vProg_Id = aProg(i)
              sGetProg vProg_Id
          %>
          <tr>
            <td width="100">&nbsp;</td>
            <td class="c1" nowrap><a href="javascript:toggle('Div_<%=vProg_Id%>');"><%=vProg_Id & " - " & vProg_Title%></a></td>
          </tr>
          <tr>
            <td width="100">&nbsp;</td>
            <td class="c2">
              <div id="Div_<%=vProg_Id%>" class="div">
                <table border="0" cellspacing="0" cellpadding="2">
                  <%
                    aMods = Split(vProg_Mods, " ")
                    For j = 0 to Ubound(aMods)
                      vMods_Id = aMods(j)
                      sGetMods vMods_Id
                  %>
                  <tr>
                    <td class="c2">&nbsp; <input type="checkbox" name="vModId" value="<%=vMods_Id%>"></td>
                    <td class="c2" nowrap><%=vMods_Id & " - " & vMods_Title%></td>
                  </tr>
                  <%
                      Next
                  %>
                </table>
              </div>
            </td>
          </tr>
          <%    
            Next
          %>
        </table>
        </td>
        </td>
      </tr>
      <tr valign="top">
        <td align="center"><p class="c2">When you have selected the modules you wish to email, click <b>Go</b></p><p><input type="submit" value="Go" name="bGo" class="button"></p><h2><a href="../Source/RC_Report.asp">My Email Report</a></h2><p>&nbsp;</p></td>
      </tr>
    </form>
  </table>
  <% 
	  Server.Execute vShellLo
	%>

</body>

</html>
