<%
  '...this creates a session variable containing the actual path to the MultiUserManual folder within Repository/Documents
  '...we need the actual path for the file system object
  Dim vFolder
  vFolder = Server.MapPath(Request.ServerVariables("PATH_INFO"))
  vFolder = Left(vFolder, InstrRev(vFolder, "\"))
  Session("MultiUserManual") = vFolder
  svMultiUserManual          = vFolder
%>  