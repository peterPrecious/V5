<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Meta.asp"-->


<%
  '...extract skill set from the goals field, update new field and create new meta table of skill sets
  Function fSkillSet
    i = Instr(Ucase(vMods_Goals), "<BR><BR>") 
    If i > 0 Then
      
      j = Instr(i, Ucase(vMods_Goals), ":") 

      fSkillSet      = fUnquote(Ucase(Trim(Mid(vMods_Goals, j + 1))))
      fSkillSet      = Replace (fSkillSet, ", ", "::")
      fSkillSet      = Replace (fSkillSet, ".", "")

      '...update skillset field and edit goals field in mods table
      vSql = "UPDATE Mods SET Mods_SkillSet = '" & fSkillSet  & "', Mods_Goals = '" & fUnquote(Left(vMods_Goals, i - 1)) & "' WHERE Mods_Id = '" & vMods_Id & "' "
'     vSql = "UPDATE Mods SET Mods_SkillSet = '" & fSkillSet  & "' WHERE Mods_Id = '" & vMods_Id & "' "

      sOpenDbBase2
      oDbBase2.Execute(vSQL)
      sCloseDbBase2     

      '...update keyword field in meta table
      Dim aSkillSet
      aSkillSet = Split(fSkillSet, "::")

      For j = 0 To Ubound(aSkillSet)
        vMeta_Id = aSkillSet(j)

        sOpenDbBase2
   
        vSql = "SELECT * FROM Meta WHERE Meta_Id = '" & vMeta_Id & "'"
        Set oRsBase2 = oDbBase2.Execute(vSql)    
        If oRsBase2.Eof Then 
          vMeta_Eof = True
        Else
          vMeta_ModIds = oRsBase2("Meta_ModIds")
        End If
        Set oRsBase2 = Nothing      
    
        '...only update meta if no mods id for specific keyword
        If vMeta_Eof Or Instr(vMeta_ModIds, Ucase(vMods_Id)) = 0 Then
          '...insert record 
          vSql = "INSERT INTO Meta (Meta_Id, Meta_ModIds) VALUES ('" & vMeta_Id & "', '" & Ucase(vMods_Id) & "')"
      '   sDebug
          On Error Resume Next 
          Set oRsBase2 = oDbBase2.Execute(vSql)
          '...if unable to insert, then update
          If Not (Err.Number = 0 Or Err.Number = "") Then 
            On Error GoTo 0          
            vSql = "UPDATE Meta SET Meta_ModIds = '" & vMeta_ModIds & " " & Ucase(vMods_Id) & "' WHERE Meta_Id = '" & vMeta_Id & "'" 
      '     sDebug
            oDbBase.Execute(vSql)
          End If
          On Error GoTo 0          
    
        End If  
    
        sCloseDbBase2
        
      Next

    Else
      fSkillSet = ""    
    End If
  End Function
%>

<html>
  <head>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
    <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
    <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
    <title></title>
  </head>

  <body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
    
  <table width="100%" border="1" cellpadding="2" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
    <tr>
      <th nowrap height="30" align="left" bgcolor="#DDEEF9">Module Details</th>
      <th nowrap height="30" align="left" bgcolor="#DDEEF9">Keywords: SkillSet</th>
    </tr>
    <%

    '...read Mods
    sOpenDbBase
'   vSql = "Select TOP 50 * FROM Mods Where Right(Mods_Id, 2) = 'EN'"
    vSql = "Select * FROM Mods"
    Set oRsBase = oDbBase.Execute(vSQL)    
    Do While Not oRsBase.Eof 
      sReadMods
      If Len(fSkillSet) > 0 Then
  %>
    <tr>
      <td valign="top"><b><%=vMods_Id%></b></td>
      <td valign="top"><%=fSkillSet%></td>
    </tr>
    <%  
      ElseIf Len(vMods_SkillSet) > 0 Then
  %>
    <tr>
      <td valign="top"><b><%=vMods_Id%></b></td>
      <td valign="top"><%=vMods_SkillSet%></td>
    </tr>
    <%  
      End If

      oRsBase.MoveNext
    Loop
    Set oRsBase = Nothing
    sCloseDbBase    
	%>
  </table>

  <% Server.Execute vShellLo %>

  </body>
</html>
