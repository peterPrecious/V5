<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  Response.write fProgXml
stop
  '...Get Mods Length
  Function fProgXml ()
    fProgXml = ""
		vSql = "SELECT Prog.Prog_Id, Prog.Prog_Title, Mods.Mods_ID, Mods.Mods_Title, Mods.Mods_Url " _
         & "FROM Prog LEFT OUTER JOIN " _
         & "Mods ON CHARINDEX(Mods.Mods_ID, Prog.Prog_Mods) > 0 " _
         & "WHERE Prog.Prog_Id = 'P1001EN' " _
         & "ORDER BY Prog.Prog_Id FOR XML AUTO " 
    sOpenDbBase    
    Set oRsBase = oDbBase.Execute(vSql)
    If Not oRsBase.Eof Then 
      fProgXml = oRsBase.Fields(0).Value
    End If
    Set oRsBase = Nothing
    sCloseDbBase    
  End Function
%>


