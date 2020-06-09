<!-- Note this upload routine can go into any web folder with iusr modify rights - no changes are required within script. -->

<html>
<head>
<title>Please Upload Your File</title>
<link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
</head>

<body>
  <center>
  <%
    Set oUp = Server.CreateObject("SoftArtisans.FileUp")
    If oUp.Form.Count > 0 Then  
      Dim vFileName  

      oUp.MaxBytes         = 0               '...no limit on individual file size
      On Error Resume Next 
      vFileName = oUp.UserFilename
      vFileName = Mid(vFileName, InstrRev(vFileName, "\") + 1)
      oUp.SaveInVirtual vFileName 
      If Err = 0 Then 
        Response.Write "<br><br>Thank you. '" & vFileName & "' was uploaded successfully. [ <a href='Default.asp'>View</a> ] [ <a href='Upload.asp'>Upload Another File</a> ]<br><br>"
      Else
        Response.Write "<p align='center'>Your file could not be uploaded because:<br><br>" & Err.Description & ".</p>"
      End If
      On Error Goto 0      
      Set oUp = Nothing
    Else
  %>
  <form enctype="multipart/form-data" method="post" action="Upload.asp">
    <h1><br>File Upload</h1><br>
    First, click <b>Browse</b> to find your local file...<br><br>
    <input type="file" name="vFile" size="59" class="c2"><br><br><br>
    ...then click <b>Submit</b> to upload it.<br><br>
    <input type="submit" value="Submit" class="c2">
  </form>  
  <% 
    End If 
  %>  
  </center>
</body>

</html>

