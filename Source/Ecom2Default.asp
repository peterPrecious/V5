<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Querystring.asp"-->

<% 
    Dim vLeft, vRight, vMode
   
    sGetQueryString '...this is just used to grab the vTraining field in for ERGP who wants to sell group2 content and position the visitor at the appropriate group
    
    Session("Ecom_Media") = Request("vEcom_Media")

    If fNoValue(Session("Ecom_Media")) Then 
      Response.Redirect "EcomError.asp?vMsg=" & Server.UrlEncode("No media selected in " & Request.ServerVariables("Script_Name"))
    End If  

    If Session("Ecom_Media")  = "Group" Or Session("Ecom_Media")  = "Group2" Or Session("Ecom_Media")  = "AddOn2" Then 
      Session("Ecom_Quantity") = 5
    Else
      Session("Ecom_Quantity") = 1
    End If

    If Session("Ecom_Media") = "Group2" Or Session("Ecom_Media") = "AddOn2" Then
      vLeft  = "Ecom2Catalogue.asp"
      vRight = "Ecom3Programs.asp"
    '...this works for online and group
    Else
      vLeft  = "Ecom2Catalogue.asp"
      vRight = "Ecom2Programs.asp"
    End If

    '...determine what initial category appears on right side (initiated by the left side)
    If Len(vTraining) > 0 Then 
      vLeft = vLeft & "?vInitCatlNo=" & vTraining
    End If  

%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <title>Ecom Default</title>
</head>

<frameset cols="35%,*" framespacing="0" border="0" frameborder="0">
  <frame name="Left" src="<%=vLeft%>" target="Right" scrolling="auto">
  <frame name="Right" src="<%=vRight%>" target="_self" scrolling="auto">
</frameset>

</html>
