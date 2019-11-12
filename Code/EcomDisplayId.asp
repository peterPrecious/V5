<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/QueryString.asp"-->

<%
  Dim vNoSource, bBypassDisplay
  vSource = Request("vSource")
  If Len(vSource) = 0 Then vNoSource = True Else vNoSource = False
  vCust_Id = Session("EcomCust")
  bBypassDisplay = fDefault(Session("Ecom_BypassDisplay"), False)
%>

<html>

<head>
  <title>EcomDisplayId</title>
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

      <% If bBypassDisplay Then %>

      <tr>
        <td style="text-align:center">
          <h1><!--webbot bot='PurpleText' PREVIEW='Thank you'--><%=fPhra(000246)%>.</h1>
          <h2>The programs just ordered are now available.</h2>
          <h2><!--webbot bot='PurpleText' PREVIEW='Please click <b>Continue</b>'--><%=fPhra(000351)%></h2>
          <p><input type="button" onclick="location.href='<%=vSource%>'" value="<%=bContinue%>" name="B3" class="button"></p>
        </td>
      </tr>        

      <% ElseIf Session("PassThru") Then %>

      <tr>
        <td style="text-align:center">
          <h1><!--webbot bot='PurpleText' PREVIEW='Thank you'--><%=fPhra(000246)%>.</h1>
          <h2>The programs just ordered are now available.</h2>
          <h2><!--webbot bot='PurpleText' PREVIEW='Please click <b>Continue</b>'--><%=fPhra(000351)%></h2>
          <p><input type="button" onclick="location.href='<%=vSource%>'" value="<%=bContinue%>" name="B3" class="button"></p>
          <h2 style="text-align:center"><!--webbot bot='PurpleText' PREVIEW='or you can click <b>Sign In</b> to begin now.'--><%=fPhra(000352)%></h2>
          <p style="text-align:center"><input onclick="location.href='//<%=svHost%>/default.asp?vCust=<%=Session("EcomCust")%>&amp;vId=<%=Session("EcomId")%>'" type="button" value="<%=bSignIn%>" name="B4" class="button"></p>
          <h2><!--webbot bot='PurpleText' PREVIEW='For your reference, this is your Customer Id and Password.'--><%=fPhra(000567)%></h2>
          <div style="text-align:center">
            <table class="table">
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
        </td>
      </tr>

      <% Else %>

      <tr>
        <td style="text-align:center">
          <h1><!--webbot bot='PurpleText' PREVIEW='Thank you'--><%=fPhra(000246)%>.</h1>
          <h2><!--webbot bot='PurpleText' PREVIEW='Here is your Customer Id and new Password.'--><%=fPhra(000353)%></h2>
          <table style="width:300px; margin:auto;">
            <tr>
              <th style="width:50%"><!--webbot bot='PurpleText' PREVIEW='Customer Id'--><%=fPhra(000111)%> : </th>
              <td><%=Session("EcomCust")%></td>
            </tr>
            <tr>
              <th><!--webbot bot='PurpleText' PREVIEW='Password'--><%=fPhra(000211)%> : </th>
              <td><%=Session("EcomId")%></td>
            </tr>
          </table>
          <h2><!--webbot bot='PurpleText' PREVIEW='To begin click <b>Continue</b> where you can enter above Password.'--><%=fPhra(000354)%></h2>
          <p><input type="button" onclick="location.href='<%=vSource%>'" value="<%=bContinue%>" name="B5" class="button"></p>
          <h2 style="text-align:center"><!--webbot bot='PurpleText' PREVIEW='You can also click below to automatically <b>Sign In</b>.'--><%=fPhra(000345)%></h2>
          <p style="text-align:center"><input onclick="location.href='//<%=svHost%>/default.asp?vCust=<%=Session("EcomCust")%>&amp;vId=<%=Session("EcomId")%>'" type="button" value="<%=bSignIn%>" name="B6" class="button"></p>
        </td>

      </tr>

      <% End If%>


      <% If Not bBypassDisplay Then %>
      <tr>
        <td style="text-align:center"><br />
          <table style="width:300px; margin:auto">
            <tr>
              <td class="rowShade"><a <%=fstatx%> href="javascript:window.print();"><img border="0" src="../Images/Icons/Printer.gif"></a></td>
              <td class="rowShade" style="text-align:left;"><!--webbot bot='PurpleText' PREVIEW='Print this page for your records.'--><%=fPhra(000346)%></td>
            </tr>
            <tr>
              <td class="rowShade"><img border="0" src="../Images/Icons/Bang.gif"></td>
              <td class="rowShade" style="text-align:left;"><!--webbot bot='PurpleText' PREVIEW='Remember to <b>Sign Off</b> after every session.'--><%=fPhra(000347)%></td>
            </tr>
            <tr>
              <td class="rowShade"><a <%=fstatx%> href="mailto:<%= fIf(Len(svCustEmail) > 0, svCustEmail, "support@vubiz.com")%>?subject=Ecommerce Issue"><img border="0" src="../Images/Icons/Email3.gif"></a></td>
              <td class="rowShade" style="text-align:left;"><!--webbot bot='PurpleText' PREVIEW='Feel free to email us if you have any questions.'--><%=fPhra(000276)%></td>
            </tr>
          </table>
        </td>
      </tr>
      <% End If %>


    </table>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


