<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="<%=svDomain%>/Inc/Vubiz.css" type="text/css" rel="stylesheet">

  <title>:: Clone Certs</title>
</head>

<body topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% 
	Server.Execute vShellHi

  Dim vMsg

  vMsg = ""
  
  
  '   This will clone the P1258EN_template and P1258FR_template from the 
  '   Virtual Repository folder to the new EN and FR values that were input
  '   vProg is entered as "P1234" which will create P1234EN and P1234FR 
  '   (even if the FR is not required)

  If Len(Request("vProg")) = 5 And Left(Request("vProg"), 1) = "P"  And IsNumeric(Right(Request("vProg"), 4)) And Lcase(Request("vPassword")) = "smoker" Then 
    sCloneCertificates Request("vProg")
    
  End If


  Sub sCloneCertificates (vProg)

    Dim oFs, vRoot, vWebs, vFolder, vFolderNew
    Set oFs = CreateObject("Scripting.FileSystemObject")

    vRoot   = Server.MapPath("\V5") 
    vMsg = vMsg & "Root:&nbsp;&nbsp;" & vRoot   & " - " & oFs.FolderExists(vRoot)

    vWebs   = Left(vRoot, Len(vRoot) - 10)
    vMsg = vMsg & "<br>Webs:&nbsp;&nbsp;" & vWebs   & " - " & oFs.FolderExists(vWebs)

    vFolder = vWebs & "\Virtual\Repository\V5_Vubz\P1258EN_template" 
    vMsg = vMsg & "<br>Folder:&nbsp;&nbsp;" & vFolder & " - " & oFs.FolderExists(vFolder)
    
    If oFs.FolderExists(vFolder) Then
      vFolderNew = vWebs & "\Virtual\Repository\V5_Vubz\" & vProg & "EN"
      oFs.CopyFolder vFolder, vFolderNew
      vMsg = vMsg & "<br>FolderNew:&nbsp;&nbsp;" & vFolderNew & " - " & oFs.FolderExists(vFolderNew)
    End If


    vFolder = vWebs & "\Virtual\Repository\V5_Vubz\P1258FR_template" 
    vMsg = vMsg & "<br>Folder:&nbsp;&nbsp;" & vFolder & " - " & oFs.FolderExists(vFolder)
    
    If oFs.FolderExists(vFolder) Then
      vFolderNew = vWebs & "\Virtual\Repository\V5_Vubz\" & vProg & "FR"
      oFs.CopyFolder vFolder, vFolderNew
      vMsg = vMsg & "<br>FolderNew:&nbsp;&nbsp;" & vFolderNew & " - " & oFs.FolderExists(vFolderNew)
    End If

    Set oFs = Nothing
  End Sub 
  
%>
  <form method="POST" action="CloneCerts.asp">
    <table border="0" width="100%" style="border-collapse: collapse" bordercolor="#DDEEF9" cellpadding="0" id="table4" cellspacing="8">
      <tr>
        <th nowrap colspan="2" valign="bottom">
        <% If Len(vMsg) > 0 Then %>
        <table border="0" id="table5" cellspacing="0" cellpadding="3" bgcolor="#DDEEF9">
          <tr>
            <td><%=vMsg%></td>
          </tr>
        </table>
        <% End If %>
        <h1>Clone a CCOHS Certificate</h1></th>
      </tr>
      <tr>
        <th align="right" nowrap width="50%">Password :</th>
        <td nowrap width="50%"><input type="text" name="vPassword" size="20"></td>
      </tr>
      <tr>
        <th align="right" nowrap width="50%">Enter New Program:</th>
        <td nowrap width="50%"><input type="text" name="vProg" size="7" maxlength="5"> ie &quot;P1234&quot;, which will create P1234EN and P1234FR</td>
      </tr>
      <tr>
        <th nowrap colspan="2" valign="top">&nbsp;<p><input type="submit" value="Clone" name="bClone" class="button"></th>
      </tr>
    </table>
  </form>
  <% 
 
  Server.Execute vShellLo
%>

</body>

</html>
