<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Repository/Documents/EcomDocumentRoutines.asp"-->
<% If Len(Request("vDelete")) > 0 Then sDeleteDocument(Request("vDelete")) %>


<html>

<head>
  <title>UploadMultiUserManual</title>
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
</head>

<body>

  <% 
    Server.Execute vShellHi 

    Dim oUp, vFileName, vDocUrl
    Set oUp = Server.CreateObject("SoftArtisans.FileUp")
    If oUp.Form.Count > 0 Then 
      Server.ScriptTimeout = 60 * 60         '...allow 60 minutes
      oUp.MaxBytes         = 0               '...no limit on individual file size
      vFileName = oUp.UserFilename
      vFileName = Mid(vFileName, InstrRev(vFileName, "\") + 1)    '...this seems to be needed for some systems to ensure we don't get full path, we just want the file name as we are saving it in the current virtual folder
      Response.Write "<br><h1>You requested your document to be uploaded as " & vFileName & ".</h1>"
      oUp.SaveInVirtual vFileName
      If Err = 0 Then 
        Response.Write "<h1>Thank you, your document was uploaded successfully with " & oUp.TotalBytes & " bytes.<br><br><br><br><a href='UploadMultiUserManual.asp'>Return</a><br><br></h1>"
      Else
        Response.Write "<p align='center'>Your file could not be uploaded because:<br><br>" & Err.Description & ".</p>"
      End If
      On Error Goto 0
      Set oUp = Nothing
    Else
  %>

  <form enctype="multipart/form-data" method="post" action="UploadMultiUserManual.asp" style="text-align: center;">
    <h1><br>Update Facilitator Manuals (in PDF format)</h1>
    <br />
    <h2>First, click <b>Browse</b> to find your local document with a file name such as <br>&quot;ACCT1234_MultiUserManual_EN.pdf&quot; or &quot;VUBZ_MultiUserManual_FR.pdf&quot;</h2>
    <br />
    <h3>NOTE: If you upload over / replace an existing Manual - no warning will be given.<br /><br />
      <input type="file" name="vFile" class="button"><br /><br />
      Then click <b>Submit</b> to upload your file.<br /><br />
    </h3>
    <input type="submit" value="Submit" class="button">
  </form>
  <br /><br />
  <h2>Documents on file...</h2>

  <div style="width: 400px; margin: auto;">
    <%=fListDocuments ("MultiUserManual")%>
  </div>

  <div style="text-align: center">
    <br /><br />
    <h3>This account (<%=svCustId%>) will offer this Document...</h3>
    <%
      vDocUrl   = fGetDocument ("MultiUserManual")
      vFileName = Mid(vDocUrl, InstrRev(vDocUrl, "/") +1)
      If vDocUrl <> "" Then
        Response.Write "<a target='_blank' href='" & vDocUrl & "'>" & vFileName& "</a>"
      Else
        Response.Write "No documents are available for this Account."
      End If
    %>
  </div>

  <% 
    End If 
  %>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>

