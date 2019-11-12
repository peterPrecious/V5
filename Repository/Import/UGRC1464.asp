<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<html>

<head>
  <title>UGRC1464</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5//Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/jQuery.js"></script>
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <!--#include virtual = "V5/Inc/Shell_HiLite.asp"-->

  <% 

    Dim oUp, oFs, vFileName, vSize

    Set oUp = Server.CreateObject("SoftArtisans.FileUp")
    Set oFs = Server.CreateObject("Scripting.FileSystemObject")   

    If oUp.Form.Count > 0 Then
  
      On Error Resume Next 
      vFileName = oUp.UserFilename

      If Len(vFileName) = 0 Then
        Response.Redirect "Message.asp?vNext=Default.asp&vMsg=" & Replace("No file has been selected.", " ", "+")
      End If

      vFileName = Ucase(Mid(vFileName, InstrRev(vFileName, "\") + 1))
      If vFileName <> "UGRC1464.TXT" Then
        Response.Redirect "Message.asp?vNext=Default.asp&vMsg=" & Replace("Please browse to find the correct file.", " ", "+")
      End If
      
      vSize = oUp.TotalBytes

      '...delete file if it exists (Server 2008/IIS7 issue)
      If oFs.FileExists(Server.MapPath(vFileName)) Then
        oUp.Delete Server.MapPath(vFileName)
      End If

      oUp.SaveInVirtual vFileName
      Set oUp = Nothing
      If Err = 0 Then 
        If vFileName = "UGRC1464.TXT" Then 
          Response.Redirect "UGRC1464_Ok.asp"
        End If
      Else
        Response.Redirect "Message.asp?vNext=Default.asp&vMsg=" & Replace("Your file could not be uploaded because:<br><br>" & Err.Description & ".", " ", "+")
      End If
    
    Else

  %>
  <form enctype="multipart/form-data" method="post" action="UGRC1464.asp" style="width:600px; margin:auto;">
    <h1>Unified Grocers | Upload Learner Profiles</h1>
    <div class="c3">
      When ready to import the learner profiles click &quot;Browse...&quot; to find: <b>UGRC1464.txt</b>.
      <span style="color:red">Note: this must be a &quot;tab delimited file&quot; with that exact filename.</span>
      When it appears in the text box below click <b>Submit</b>.
      <br /><br />Any records imported will added to the user table if new or will overwrite previously imported user records.
      Any users that are on file but NOT contained in this tab delimited file will be flagged as inactive.
      No user records are ever deleted.<br /><br />
      First, click &quot;Browse&quot; to find the appropriate file...<br><br>
      <div style="text-align:center"><input type="file" name="vLearners"><br /></div>
      <br />Then click &quot;Submit&quot; to upload the file.
    </div>
    <div style="text-align:center">
      <input class="button" type="submit" name="sNext" id="txt04" value="Next" />
    </div>

    <h6>Note:&nbsp; this can take several minutes to upload - only press the submit button once! <br>Imported summaries will be displayed upon completion.</h6>
  </form>

  <% 
    End If 
  %>

  <!--#include virtual = "V5/Inc/Shell_LoLite.asp"-->


</body>

</html>
