﻿<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Prod.asp"-->
<!--#include virtual = "V5/Repository/Documents/EcomDocumentRoutines.asp"-->

<%
  Dim vNoOptions, vUrl, vWidth, vProdSpecials, vDocUrl

  Session("Ecom_Quantity")   = 0	'...initialize quantity
  Session("Ecom_CdDiscount") = 0	'...restore the percentage discount
  Session("ProdNo")          = 0	'...initialize basket product no					
  Session("ProdMax")         = 0	'...initialize basket no products

  '...determine what the options are
  vContentOptions = Request.Querystring("vContentOptions")
  If Len(vContentOptions) <> 4 Then
    Response.Redirect "Error.asp?vReturn=n&vErr=" & Server.UrlEncode("Invalid Action.  Please contact Vubiz Support") 
  End If  

  '...determine the number of content options?
  vNoOptions = 0
  If Left(vContentOptions, 1)   = "Y" Then vNoOptions = 1
  If Mid(vContentOptions, 2, 1) = "Y" Then vNoOptions = vNoOptions + 1
  If Mid(vContentOptions, 3, 1) = "Y" Then vNoOptions = vNoOptions + 1
  If Right(vContentOptions, 1)  = "Y" Then vNoOptions = vNoOptions + 1   '...the old Prods is being used for group2 addon2
    
  Select Case vNoOptions
    Case 1 : vWidth = 100
    Case 2 : vWidth =  50
    Case 3 : vWidth =  33
  End Select

%>
<html>

<head>
  <title>Ecom2Start</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>AC_FL_RunContent = 0;</script>
  <script src="/V5/Inc/AC_RunActiveContent.js"></script>
  <script>
    function jTitle (vTitle, vImage) {
      var vParm = "title=" + vTitle + '&image=/V5/Images/Titles/' + vImage;
      AC_FL_RunContent('codebase','//download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0','name','flashVars','width','265','height','85','align','middle','id','flashVars','src','/V5/Images/Titles/VuTitles','FlashVars',vParm,'quality','high','bgcolor','#ffffff','allowscriptaccess','sameDomain','allowfullscreen','false','pluginspage','///go/getflashplayer','movie','/V5/Images/Titles/VuTitles');
    }
  </script>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table class="table">
    <% If vNoOptions = 0 Then %>
    <tr>
      <td style="text-align: center">
        <h5>
          <!--webbot bot='PurpleText' PREVIEW='Sorry, there are no programs available'--><%=fPhra(000242)%>.</h5>
      </td>
    </tr>
    <% 
       ElseIf vNoOptions = 1 Then
         If Left(vContentOptions, 1) = "Y" Then 
           Response.Redirect "Ecom2Default.asp?vEcom_Media=Online"
         ElseIf Mid(vContentOptions, 2, 1) = "Y" Then 
           Response.Redirect "Ecom2Default.asp?vEcom_Media=Group"
         ElseIf Mid(vContentOptions, 3, 1) = "Y" Then 
           Response.Redirect "Ecom2Default.asp?vEcom_Media=Group2"
         ElseIf Mid(vContentOptions, 4, 1) = "Y" Then 
           Response.Redirect "Ecom2Default.asp?vEcom_Media=AddOn2"
         End If
       Else
    %>
    <tr>
      <td style="text-align: center">

        <%
          Dim aTools, vEcom
          If Len(svBrowser) > 0 Then
            aTools = Split(Ucase(svBrowser), "|")
            vEcom    = aTools(6) 
'           If vEcom = "N" Then
        %>

        <%
'            End If
          End If
        %>
        
        <h1><!--webbot bot='PurpleText' PREVIEW='Select your preferred learning format.'--><%=fPhra(000888)%></h1>

      </td>
    </tr>
  </table>

  <table class="table">
    <tr>
      <% 
        If Left(vContentOptions, 1) = "Y" Then 
          vUrl = "Ecom2Default.asp?vEcom_Media=Online"
      %>
      <td style="width: <%=vWidth%>%; padding: 10px; text-align: center">
<!--        
  <script>jTitle("<%=fPhraH(000376)%>", 'SingleLicense.jpg')</script>
-->
        <img src="../Images/Ecom/SingleLearnerLicense_<%=svLang %>.png" />


        <p class="c3" style="text-align: left">
        <!--webbot bot='PurpleText' PREVIEW='Use this option if you are purchasing and want immediate access (with no administrative tracking privileges) to a course or bundle of courses. This option is ideal for an individual learner.&nbsp; You will receive your password onscreen at the end of your purchase transaction. With this purchase, access is for 90 days unless stated otherwise.'--><%=fPhra(001319)%></p>

        <% If svCustId = "VUBZ5678" And svLang = "EN" Then %>  
        <input onclick="window.open('//store.vubiz.com/store')" type="button" value="<%=bContinue%>" name="bContinue" class="button">
        <% Else %>
        <input onclick="location.href = '<%=vUrl%>'" type="button" value="<%=bContinue%>" name="bContinue" class="button">
        <% End If  %>

      </td>
      <%
        End If     
      
        If Mid(vContentOptions, 2, 1) = "Y" Then 
          vUrl = "Ecom2Default.asp?vEcom_Media=Group"
      %>
      <td style="width: <%=vWidth%>%; padding: 10px; text-align: center">
<!--        
        <script>jTitle("<%=fPhraH(000377)%>", 'MultiLicense.jpg')</script>
-->
        <img src="../Images/Ecom/MultipleLearnerLicense_<%=svLang%>.png" />


        <p class="c3" style="text-align: left">
          <!--webbot bot='PurpleText' PREVIEW='If you are purchasing 5 or more learner licenses and would like administrative tracking privileges, utilize this option. This option is ideal if the company would like to view trainee progress toward completion of training modules and offers discounted pricingfor purchasing multiple licenses.<br><br>Once you have received your password and Customer ID, utilize the Facilitator Manual for step-by-step directions on how to assign learners and access reports. You will receive your Facilitator password and customer ID onscreen at the end of your purchase transaction. With this purchase, you are given access to the learning for 365 days.'--><%=fPhra(001862)%></p>

          <input onclick="location.href = '<%=vUrl%>'" type="button" value="<%=bContinue%>" name="bContinue0" class="button">
<!--                TEMPORARILY SUSPENDED-->


          <a <%=fstatx%> target="_blank" href="../Images/Documents/VuMuiltiUserLic1.doc"><br /><br />
          <!--webbot bot='PurpleText' PREVIEW='Facilitator Manual'--><%=fPhra(000870)%></a>
      </td>
      <% 
        End If 

        If Mid(vContentOptions, 3, 1) = "Y" Then
          vUrl = "Ecom2Default.asp?vEcom_Media=Group2"
          vDocUrl = fGetDocument ("MultiUserManual")
      %>
      <td style="width: <%=vWidth%>%; padding: 10px; text-align: center">

        <img src="../Images/Ecom/MultipleLearnerLicense_<%=svLang%>.png" />
<!--
        <script>jTitle("<%=fPhraH(000377)%>", 'MultiLicense.jpg')</script>
-->

        <p class="c3" style="text-align: left">
          <!--webbot bot='PurpleText' PREVIEW='If you are purchasing 5 or more learner licenses and would like administrative tracking privileges, utilize this option. This option is ideal if the company would like to view trainee progress toward completion of training modules and offers discounted pricing for purchasing multiple licenses.<br><br>Once you have received your password and Customer ID, utilize the Facilitator Manual for step-by-step directions on how to assign learners and access reports. You will receive your Facilitator password and customer ID onscreen at the end of your purchase transaction. With this purchase, you are given access to the learning for 365 days.'--><%=fPhra(000954)%></p>
        <input onclick="location.href = '<%=vUrl%>'" type="button" value="<%=bContinue%>" name="bContinue3" class="button">
<!--        TEMPORARILY SUSPENDED-->

        <% If vDocUrl <> "" Then %><br /><br /><a <%=fstatx%> target="_blank" href="<%=vDocUrl%>">
          <!--webbot bot='PurpleText' PREVIEW='Facilitator Manual'--><%=fPhra(000870)%></a><% End If%>
      </td>
      <% 
        End If 

        If Mid(vContentOptions, 4, 1) = "Y" Then
          vUrl = "Ecom2Default.asp?vEcom_Media=AddOn2"
          vDocUrl = fGetDocument ("MultiUserManual")
      %>
      <td style="width: <%=vWidth%>%; padding: 10px; text-align: center">

        <img src="../Images/Ecom/MultipleLearnerLicense_<%=svLang%>.png" />
<!--
       <script>jTitle("<%=fPhraH(000377)%>", 'MultiLicense.jpg')</script>
-->


        <p class="c3" style="text-align: left">
          <!--webbot bot='PurpleText' PREVIEW='If you are purchasing 5 or more learner licenses and would like administrative tracking privileges, utilize this option. This option is ideal if the company would like to view trainee progress toward completion of training modules and offers discounted pricing for purchasing multiple licenses.<br><br>Once you have received your password and Customer ID, utilize the Facilitator Manual for step-by-step directions on how to assign learners and access reports. You will receive your Facilitator password and customer ID onscreen at the end of your purchase transaction. With this purchase, you are given access to the learning for 365 days.'--><%=fPhra(000954)%></p>
        <input onclick="location.href = '<%=vUrl%>'" type="button" value="<%=bContinue%>" name="bContinue3" class="button">
<!--                TEMPORARILY SUSPENDED-->

        <% If vDocUrl <> "" Then %><br /><br /><a <%=fstatx%> target="_blank" href="<%=vDocUrl%>">
          <!--webbot bot='PurpleText' PREVIEW='Facilitator Manual'--><%=fPhra(000870)%></a><% End If%>
      </td>
      <%  
        End If 
      %>
    </tr>

    <% 
      End If 
    %>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


