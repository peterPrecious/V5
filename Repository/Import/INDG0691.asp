<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="<%=svDomain%>/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <link href="/V5/Inc/<%=Left(svCustId, 4)%>.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
</head>


<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% 
    Server.Execute vShellHi 

    Dim oUp, vFileName, vSize
    Set oUp = Server.CreateObject("SoftArtisans.FileUp")

    If oUp.Form.Count > 0 Then
  
      On Error Resume Next 
      vFileName = oUp.UserFilename

      If Len(vFileName) = 0 Then
        Response.Redirect "Message.asp?vNext=Default.asp&vMsg=" & Replace("No file has been selected.", " ", "+")
      End If

      vFileName = Ucase(Mid(vFileName, InstrRev(vFileName, "\") + 1))
      If vFileName <> "INDG0691.TXT" Then
        Response.Redirect "Message.asp?vNext=Default.asp&vMsg=" & Replace("Please browse to find the correct file.", " ", "+")
      End If
      
      vSize = oUp.TotalBytes
      oUp.SaveInVirtual vFileName
      Set oUp = Nothing
      If Err = 0 Then 
        If vFileName = "INDG0691.TXT" Then 
          Response.Redirect "INDG0691_Ok.asp"
        End If
      Else
        Response.Redirect "Message.asp?vNext=Default.asp&vMsg=" & Replace("Your file could not be uploaded because:<br><br>" & Err.Description & ".", " ", "+")
      End If
    
    Else

  %>

    <form enctype="multipart/form-data" method="post" action="INDG0691.asp">
      <div align="center">
      <table border="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9" width="80%">
        <tr>
          <td align="center"><h1>Chapter Indigo | Upload Learner Profiles</h1>
          <p class="c2" align="left">When ready to import the new Learner file click &quot;Browse...&quot; to find: <b>INDG0691.txt</b>.&nbsp; <font color="#FF0000">Note: this must be a &quot;tab delimited file&quot; with that exact filename.&nbsp; </font>When it appears in the text box below click <b>Submit</b>.&nbsp; Any records imported will added to the user table if new or will overwrite previously imported user records.&nbsp; Any users that are on file but NOT contained in this tab delimited file will be flagged as inactive.&nbsp; No user records are ever deleted.&nbsp; </p><p class="c2">First, click &quot;Browse&quot; to find the appropriate file...<br><br><input type="file" name="vLearners" size="59"></p>
          <h2 align="center">Then click &quot;Submit&quot; to upload the file.<br><br><input type="submit" value="Submit"></h2>
          <h5 align="center">Note:&nbsp; this takes about 5 minutes to upload - only press the submit button once! <br>Imported summaries will be displayed upon completion.</h5>
          </td>
        </tr>
        </table>
    </form>
  
  <% 
    End If 
  %>

  <!--#include virtual = "V5/Inc/Shell_LoLite.asp"-->


</body>

</html>