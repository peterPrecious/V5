<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="http://vubiz.com/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <title>Upload</title>
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
      Response.Redirect "Message.asp?vNext=Import.asp&vMsg=" & Replace("No file has been selected.", " ", "+")
    End If

    vFileName = Ucase(Mid(vFileName, InstrRev(vFileName, "\") + 1))

    If vFileName <> Ucase(svCustId) & ".CSV" Then
      Response.Redirect "Message.asp?vNext=Import.asp&vMsg=" & Replace("Please browse to find the file '" & svCustId & ".csv' to upload.", " ", "+")
    End If
    
    vSize = oUp.TotalBytes
    oUp.SaveInVirtual "Import/" & vFileName
    Set oUp = Nothing
    If Err = 0 Then 
      Response.Redirect "ImportOk.asp"
    Else
      Response.Redirect "Message.asp?vNext=Import.asp&vMsg=" & Replace("Your file could not be uploaded because:<br><br>" & Err.Description & ".", " ", "+")
    End If
  
  Else
%>

<form enctype="multipart/form-data" method="post" action="Import.asp">
  <table border="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9" width="100%">
    <tr>
      <td>
        <h1 align="center">Upload Learner Profiles&nbsp; (Custom)</h1>
        <p class="c2">This program allows you to import your learner profiles into this account (<%= svCustId%>).&nbsp; Do not use this service to upload facilitators or managers, these must be done manually as they required special rights.</p>
        <blockquote>
          <ul class="c2">
            <li>Prepare a populated CSV file (comma separated values) that contains all or some of these fields: Learner Id, Learner Password (if so configured), First Name, Last Name, Email Address and Group Id.&nbsp; Note: the only mandatory fields is the User Id and Password fields.<br>&nbsp;</li>
            <li>This file MUST have the file name: <b><%= svCustId & ".csv"%></b>.&nbsp; <a target="_blank" href="CUSTOMER.CSV">click here to view a sample CSV file.</a> Note: you can save this file as <b><%= svCustId & ".csv"%></b> onto your desktop and it can be populated for importing.<br>&nbsp;</li>
            <li>Once your file contains all the learner fields, you can begin the Import process by first clicking on the <b>Browse...</b> button to find the <b><%= svCustId & ".csv"%></b> file.&nbsp; Then, click on the <b>Submit</b> button to import these values into the database.<br>&nbsp;</li>
            <li>Note:&nbsp; All learner Ids must be unique.&nbsp; If you import a learner with an ID that is already on file, the information on file will be overwritten by the new values.&nbsp; Also, if you have duplicate Ids in the CSV file, the latter record will overwrite the previous.&nbsp; It is very important that you do not overwrite records in error as original records may have logging information attached to them that would now be attached to the new records.&nbsp; Any records with missing Ids will be ignored.&nbsp; You will have a chance to view the records you update before they are imported.</li>
          </ul>
        </blockquote>
      </td>
    </tr>
    <tr>
      <td align="center"><input type="file" name="vFile" size="33" class="c2"></td>
    </tr>
    <tr>
      <th align="center" class="c2">
        First, click <b>Browse</b> to find file <b><%= svCustId & ".csv"%></b><br>
        <br>&nbsp;Then click <b>Submit </b>to upload the file.<br>
        <br>&nbsp;<input type="submit" value="Submit" class="c2"><br><br>
        Note:&nbsp; large files can take a few minutes to upload - only press the submit button once !
      </th>
    </tr>
  </table>
</form>
<% 
    End If 
%>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
