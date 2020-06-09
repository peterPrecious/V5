<%
   Dim vHost
   vHost = "s2.vubiz.com:8080/v5"    '...this is the default
'  vHost = "localhost/v5"            '...this is for testing
   Response.Redirect "//" & vHost & "/Default.asp?vCust=VUID2330&vId=FACILITATOR&vTest=Y&vBookmark=Y&vQModId=" & Request.QueryString("vModId")  
%>