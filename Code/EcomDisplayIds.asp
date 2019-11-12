<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<%
  Dim vNoSource
  If Len(svCustReturnUrl) = 0 Then vNoSource = True Else vNoSource = False
%>

<html>

<head>
  <title>EcomDisplayIds</title>
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
          <h1><!--webbot bot='PurpleText' PREVIEW='Thank you'--><%=fPhra(000246)%>.</h1>
          <h2><!--webbot bot='PurpleText' PREVIEW='Here is your new Customer Id and Password <br>which you will need to access your new service.'--><%=fPhra(000343)%></h2>

          <div align="center">
            <table border="1" cellpadding="2" cellspacing="0" bordercolor="#FF0000" id="table1">
              <tr>
                <td>

           <table style="width:300px; margin:auto;">
            <tr>
              <th style="width:50%"><!--webbot bot='PurpleText' PREVIEW='Customer Id'--><%=fPhra(000111)%> : </th>
              <td style="width:50%"><%=Session("EcomCust")%></td>
            </tr>
            <tr>
              <th><!--webbot bot='PurpleText' PREVIEW='Password'--><%=fPhra(000211)%> : </th>
              <td><%=Session("EcomId")%></td>
            </tr>
                  </table>
            
                </td>
              </tr>
            </table>
          </div>

          <h3 style="text-align:left;"><!--webbot bot='PurpleText' PREVIEW='Remember that you are the Facilitator for this service and the above Password gives you the rights to setup your learners (maximum'--><%=fPhra(000355)%> <%=Request("vEcom_Quantity")%><!--webbot bot='PurpleText' PREVIEW=') and monitor their progress.&nbsp; When a learner accesses this site, they will see the program(s) you have ordered, but when you enter, you will see also see an Administration tab containing links to add/edit learners, monitor their usage of programs and review their performance on assessments.&nbsp; Good luck!'--><%=fPhra(000383)%></h3>
  
          <h2><!--webbot bot='PurpleText' PREVIEW='To begin click <b>Continue</b> where you can enter above Customer Id and Password.'--><%=fPhra(000356)%></h2>        
          <p><input type="button" onclick="location.href='<%=svCustReturnUrl%>'" value="<%=bContinue%>" name="B2" class="button"></p>
  
          <h2 align="center"><!--webbot bot='PurpleText' PREVIEW='You can also click below to automatically <b>Sign In</b>.'--><%=fPhra(000345)%></h2>
          <p align="center"><input onclick="location.href='//<%=svHost%>/default.asp?vCust=<%=Session("EcomCust")%>&vId=<%=Session("EcomId")%>'" type="button" value="Sign In" name="B1" class="button"></p>

          <br /><br />
          <table style="width:300px; margin:auto">
            <tr>
              <td class="rowshade"><a <%=fstatx%> href="javascript:window.print();"><img border="0" src="../Images/Icons/Printer.gif"></a></td>
              <td class="rowshade" style="text-align:left;"><!--webbot bot='PurpleText' PREVIEW='Print this page for your records.'--><%=fPhra(000346)%></td>
            </tr>
            <tr>
              <td class="rowshade"><img border="0" src="../Images/Icons/Bang.gif"></td>
              <td class="rowshade" style="text-align:left;"><!--webbot bot='PurpleText' PREVIEW='Remember to <b>Sign Off</b> after every session.'--><%=fPhra(000347)%></td>
            </tr>
            <tr>
              <td class="rowshade"><a <%=fstatx%> href="mailto:<%= fIf(Len(svCustEmail) > 0, svCustEmail, "support@vubiz.com")%>?subject=Ecommerce Issue"><img border="0" src="../Images/Icons/Email3.gif"></a></td>
              <td class="rowshade" style="text-align:left;"><!--webbot bot='PurpleText' PREVIEW='Feel free to email us if you have any questions.'--><%=fPhra(000276)%></td>
            </tr>
          </table>

        </td>
      </tr>
      </table>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


