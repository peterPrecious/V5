<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

		<title></title>
	</head>

	<body>
	
		<% Server.Execute vShellHi %>

		<div align="center">
			<center>
			<table border="0" cellspacing="0" width="90%" cellpadding="0">
				<tr>
					<td width="100%"><%
		          Dim aMods, vLength
		          sOpenDbBase
		
		          '...go thru all the programs
		          vSql = "Select Prog_Id, Prog_Mods, Prog_Length FROM Prog"
		          Set oRsBase = oDbBase.Execute(vSQL)    
		          Do While Not oRsBase.Eof 
		            vProg_Id     = oRsBase("Prog_Id")
		            vProg_Length = oRsBase("Prog_Length")
		            vProg_Mods   = oRsBase("Prog_Mods")
		            vLength = 0
		
		            Response.Write " " & vProg_Id & "(" & vProg_Length & ":" 
		
		            '...ensure there are mods to check
		            If Len(vProg_Mods) >= 6 Then
		
		              aMods = Split(Trim(vProg_Mods), " ")
		
		              '...get mods Length
		              sOpenDbBase2    
		              For i = 0 to uBound(aMods)
		                vMods_Id = aMods(i)   
		                vSql = "SELECT Mods_Length FROM Mods WHERE Mods_Id= '" & vMods_Id & "'"
		                Set oRsBase2 = oDbBase2.Execute(vSql)
		                '...build up new length
		                If Not oRsBase2.Eof Then 
		                  vLength = vLength + oRsBase2("Mods_Length")
		                End If
		              Next    
		              sCloseDbBase2           
		
		              '...different program length?
		              If vLength <> vProg_Length Then
		                vSql = "UPDATE Prog SET Prog_Length = " & vLength & " WHERE Prog_Id = '" & vProg_Id & "'"
		                sOpenDbBase2    
		                oDbBase2.Execute(vSql)
		                sCloseDbBase2           
		
		                '...update any customers that use this program
		                sUpdateCustProgLength vProg_Id, vLength    
		                
		              End If
		
		            End If
		            Response.Write vLength & ")"
		
		            oRsBase.MoveNext
		          Loop      
		
		          Set oRsBase  = Nothing
		          Set oRsBase2 = Nothing
		
		          sCloseDbBase
		
		        %>
					<p align="center"><font color="#000080" size="1" face="Verdana"><b>All Programs and Customer &quot;Lengths&quot; have been updated.<br>
					</b>Format&nbsp;&nbsp; P9999LL(old:new)</font></p>
					<p align="center"><a href="javascript:history.back(1)"><img border="0" src="../Images/Buttons/Return_<%=svLang%>.gif"></a></p>
					<p align="center"><a href="Menu.asp"><img border="0" src="../Images/Icons/Administration.gif" alt="Click here for the Menu"></a> </p>
					</td>
				</tr>
			</table>
			</center></div>
		
		<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
	
	</body>

</html>
