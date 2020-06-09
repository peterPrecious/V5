<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  '...This is launched via an email alert - may or may not still work.  place in the root

  Session("HostDb")          = "V5_Vubz"  '...define the DB
  Session("Ecom_CDdiscount") = 25         '...store the percentage discount
  Session("Ecom_Source")     = "2274"     '...store the percentage discount
%>

<html><head><title>Vubiz Ecommerce CD Promotion</title></head>
<frameset border="0" frameSpacing="0" rows="80,*" frameBorder="0">
  <frame name="tabs" marginWidth="0" marginHeight="0" src="../Images/root/TabsPublic.asp?vTab=" scrolling="no" target="main">
  <frame name="main" marginWidth="0" marginHeight="0" src="../Images/root/EcomNoUsers.asp?vEcom_Media=CDs&vMode=More" noResize target="_self">
</frameset>
</html>