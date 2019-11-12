<%@ codepage=65001 %>

<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<html>

<head>
	<title>RTE_ProgDesc</title>
	<link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
	<script src="/V5/Inc/Functions.js"></script>
</head>

<body>

	<%
    Dim vProgId, aMods
		vProgId = (Request("vProgId")) 
    sGetProg vProgId
  %>

		<table class="table">
			<tr>
				<td>
					<h1><br><%=vProg_Title %></h1>
					<h2><!--webbot bot='PurpleText' PREVIEW='Program Id'--><%=fPhra(001446)%>: <%=vProg_Id %></h2>
					<%=vProg_Desc%>

          <h3><!--webbot bot='PurpleText' PREVIEW='Estimated program length'--><%=fPhra(000130)%> :</h3>
					<%=f5%><b><%=vProg_Length%></b>&ensp;<!--webbot bot='PurpleText' PREVIEW='Hour(s)'--><%=fPhra(000143)%>.
	
					<% 
						If Len(Trim(vProg_Mods)) > 0 Then 
							Response.Write "<h3>" & fPhraH(001448) & " :</h3>"
							aMods = Split(Trim(vProg_Mods), " ")
							For i = 0 To Ubound(aMods)
								sGetMods(aMods(i))
								If vMods_Active Then Response.Write f5 & "<b>" + vMods_Id & "</b>&ensp;" & vMods_Title & "<br>"
							Next
						End If

						If Len(Trim(vProg_Assessment)) > 0 Then 
							Response.Write "<h3>" & fPhraH(001449) & " :</h3>"
							sGetMods(Trim(vProg_Assessment))
							If vMods_Active Then Response.Write f5 & "<b>" + vMods_Id & "</b>&ensp;" & vMods_Title & "<br>"
						End If
					%> 

				</td>
			</tr>
		</table>

</body>
</html>


