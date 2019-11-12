<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<html>

	<head>
		<title>VUBIZ Secure File Download</title>
		<link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
	</head>
	
	<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">
	
	<!--#include virtual = "V5/Inc/Shell_HiSolo.asp"-->

	<div align="center">
		<table border="1" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#00FFFF">
			<% 
		    If Ucase(Request.Form("vPass")) <> "VUB!Z" Then			
			%>
			<tr><td colspan="3">
				<form method="POST" action="Download.asp">
					Password: <input type="password" name="vPass" size="20"><input type="submit" value="GO" name="bGo">
				</form>
			</td></tr>
			<%			
				Else
			%>
			<tr><td class="c1" height="30" colspan="3" align="center"><b>Documents</b></td></tr>
			<tr><td class="c1" height="30">File Id</td><td class="c1" height="30">Created</td><td class="c1" height="30" align="center">Action</td></tr>
			<%			
			  Dim oFs,oFo, vFolder, vFile
		    vFolder = Server.MapPath("Default.asp")
		    vFolder = Left(vFolder, Len(vFolder) - 12)
			  Set oFs = Server.CreateObject("Scripting.FileSystemObject")
			  Set oFo = oFs.GetFolder(vFolder)
			  For Each vFile In oFo.Files		
			  	If vFile.Name <> "Default.asp" and vFile.Name <> "Download.asp" Then						
			%>
			<tr><td class="c2">
				<a href="<%=vFile.Name%>"><%=vFile.Name & f5()%></a>
			</td><td class="c2">
			 <%= vFile.DateCreated & f5()%>
			</td><td class="c2" style="font-variant: small-caps" align="center">			 
				<a href="#" onclick="alert('Future feature')">DELETE</a></td></tr>
			<%
						End If
				  Next
				  Set oFo = Nothing
				  Set oFs = Nothing			
		    End If 
		  %>
		</table>
	</div>
	<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->
	
	</body>

</html>