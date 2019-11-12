<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->


<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <link href="/V5/Inc/<%=Left(svCustId, 4)%>.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% 
    Server.Execute vShellHi 
    
    Dim oUp, oFs, vFileName, vSize

    Set oUp = Server.CreateObject("SoftArtisans.FileUp")
    Set oFs = Server.CreateObject("Scripting.FileSystemObject")   

    If oUp.Form.Count > 0 Then
  
      On Error Resume Next 
      vFileName = oUp.UserFilename

      If Len(vFileName) = 0 Then
        Response.Redirect "Message.asp?vNext=Import.asp&vMsg=" & Replace("No file has been selected.", " ", "+")
      End If

      vFileName = Lcase(Mid(vFileName, InstrRev(vFileName, "\") + 1))

      If vFileName <> "elearn.csv" And vFileName <> "department codes.csv" Then
        Response.Redirect "Message.asp?vNext=CMSS2592.asp&vMsg=" & Replace("Please browse to find the correct file.", " ", "+")
      End If
      

      '...delete file if it exists (Server 2008/IIS7 issue)
      If oFs.FileExists(Server.MapPath(vFileName)) Then
        oUp.Delete Server.MapPath(vFileName)
      End If

      vSize = oUp.TotalBytes
      oUp.SaveInVirtual vFileName
      Set oUp = Nothing
      If Err = 0 Then 
        If vFileName = "elearn.csv" Then 
          Response.Redirect "CMSS2592_Ok.asp"
        End If
      Else
        Response.Redirect "Message.asp?vNext=Import.asp&vMsg=" & Replace("Your file could not be uploaded because:<br><br>" & Err.Description & ".", " ", "+")
      End If
    
    Else

  %>

    <form enctype="multipart/form-data" method="post" action="CMSS2592.asp">
      <div align="center">
      <table border="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9" width="80%">
        <tr>
          <td align="center"><h1>City of Mississauga | Upload Learner Profiles</h1>
            <p class="c2" align="left">When ready to Upload (Import) the new Learner profiles click <b>Browse...</b> to find: <b>elearn.csv</b>.&nbsp; When it appears in the text box below click <b>Submit</b>.&nbsp; Any records uploaded will be added to the learner table if new or will overwrite previously uploaded learner records.&nbsp; Any learners that are on the learner table but NOT contained in this CSV file will be flagged as inactive.&nbsp; No learner records are ever deleted.&nbsp; Uploaded summaries will be displayed upon completion.</p><p class="c2">First, click <b>Browse</b> to find the appropriate file...<br><br>
            <input type="file" name="vLearners" size="24" class="button"></p>
            <p class="c2" align="center">Then click <b>Submit</b> to upload the file.</p>
            <a id="aNext" class="butShell" href="javascript:void(0)"><span class="butIcon butNext"></span><input class="butInput" type="submit" name="sNext" id="txt04" value="Next" /></a>
            <h5 align="center">Note:&nbsp; this can take several minutes to upload - only press the submit button once! <br>Imported summaries will be displayed upon completion.</h5>
          </td>
        </tr>
        </table>
      </div>
    </form>
  
  <% 
    End If 
  %>

  <!--#include virtual = "V5/Inc/Shell_LoLite.asp"-->

</body>

</html>