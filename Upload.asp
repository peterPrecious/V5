<!-- Note this upload routine can go into any web folder with iusr modify rights - no changes are required within script. -->

<html>
<head>
  <title>:: Please Upload Your File</title>
</head>
<body>

  <%
    Dim oUp, oFs
  
    Set oFs = Server.CreateObject("Scripting.FileSystemObject")   
    Set oUp = Server.CreateObject("SoftArtisans.FileUp")

    If oUp.Form.Count > 0 Then
  
      Dim vFileName
  
      Server.ScriptTimeout = 60 * 20
      oUp.MaxBytes         = 0               '...no limit on individual file size
'     oUp.MaxBytesToCancel = 1000000000      '...1gig total transfer max
  
      On Error Resume Next 
      vFileName = oUp.UserFilename
      vFileName = Mid(vFileName, InstrRev(vFileName, "\") + 1)

      '...delete file if it exists (Server 2008/IIS7 issue)
      If oFs.FileExists(Server.MapPath(vFileName)) Then
        oUp.Delete Server.MapPath(vFileName)
      End If

      oUp.SaveInVirtual vFileName
  
      If Err = 0 Then 
        Response.Write "<p align='center'>Thank you, your file was uploaded successfully with " & oUp.TotalBytes & " bytes.</p>"
      Else
        Response.Write "<p align='center'>Your file could not be uploaded because:<br><br>" & Err.Description & ".</p>"
      End If
      On Error Goto 0
      
      Set oUp = Nothing
    Else
  %>

  <form enctype="multipart/form-data" method="post" action="Upload.asp">
    <p align="center"><br><br>First, click &quot;Browse&quot; to find your local file...<br>&nbsp;
    <input type="file" name="vFile" size="59">
    <br><br><br><br>Then click &quot;Submit&quot; to upload your file.&nbsp; <br>Note:&nbsp; large files take longer to upload, only press the submit button once ! </p>
    <p align="center"><input type="submit" value="Submit"></p>
  </form>

  <% 
    End If 
  %>
</body>
</html>
