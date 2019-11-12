<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<html>

	<head>
		<title>VUBIZ Secure File Upload</title>
		<link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
	</head>
	
	<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">
	
		<!--#include virtual = "V5/Inc/Shell_HiSolo.asp"-->
		
		<% 
			Dim oUp, vAccessId
	    Set oUp = Server.CreateObject("SoftArtisans.FileUp")
	    vAccessId = Request("vAccessId")
	
			'... service requires any 8 char id (typicall cust id like VUBZ2274 - it bypasses security
	
	    If oUp.Form.Count > 0 Then  
	      Dim vFileName  
	      Server.ScriptTimeout = 60 * 60         '...allow 60 minutes
	      oUp.MaxBytes         = 0               '...no limit on individual file size
	      vAccessId = Ucase(oUp.Form("vAccessId"))

		    If Len(vAccessId) <> 8 Then
		      Response.Write "<p align='center' class='c5'><br><br>That is not a valid Access Id!<br><br><br></p>"

		    Else		      

		      vFileName = oUp.UserFilename
		      vFileName = vAccessId & "_" & Mid(vFileName, InstrRev(vFileName, "\") + 1)	
		      On Error Resume Next 
		      oUp.SaveInVirtual vFileName 
		      If Err = 0 Then 
		        Response.Write "<p align='center'>Thank you. Your file was uploaded successfully</pr>"
		      Else
		        Response.Write "<p align='center'>Your file could not be uploaded because:<br><br>" & Err.Description & ".</p>"
		      End If
		      On Error Goto 0      
		      Set oUp = Nothing

				End If

	    Else
			%>
			<form enctype="multipart/form-data" method="post" action="Default.asp">
				<div align="center">
					<table border="0" cellspacing="0" cellpadding="0">
						<tr>
							<td align="center" colspan="2"><img border="0" src="/v5/images/Logos/vubz.jpg" width="146" height="50"><h1>VUBIZ Generic File Upload</h1>
							<p align="left">Notes:</p>
							<ol style="text-align: left">
								<li>This service enables you to upload any file to the Vubiz servers.</li>
								<li>For security uploading enter entered this site using <a href="https://learn.vubiz.com">https://learn.vubiz.com</a>.</li>
								<li>If the file you upload already exists, it will be overwritten.</li>
								<li>To begin, enter your Access Id then click <b>Browse</b> to find your local file...</li>
								<li>Then click <b>Submit</b> to upload it.</li>
							</ol>
							</td>
						</tr>
						<tr>
							<td align="right">Access ID :  </td>
							<td align="left"><input type="text" name="vAccessId" size="20"></td>
						</tr>
						<tr>
							<td align="right">Next :  </td>
							<td align="left"><input type="file" name="vFile" size="30"></td>
						</tr>
						<tr>
							<td align="center" colspan="2"><br>
							<input type="submit" value="Submit" class="button"></td>
						</tr>
					</table>
				</div>
			</form>
			<% 
		    End If 
		  %>
		<!--#include virtual = "V5/Inc/Shell_Lo.asp"-->	
	</body>

</html>
