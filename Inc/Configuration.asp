<%
  '...define site info and database info
  Session("Site")        = "VuImport"
  Session("Domain")      = Lcase(Request.ServerVariables("HTTP_HOST")) 
  Session("Host")        = Session("Domain") & "/" & Lcase(Session("Site"))
  Session("Secure")      = False

  '...define sql default db and sa password
  Session("HostDb")      = "V5_Vubz"
  Session("HostDbPwd")   = "vudb2112mississauga"
%>  