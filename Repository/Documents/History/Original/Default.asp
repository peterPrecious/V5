<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  Const ForReading = 1, ForWriting = 2
  Dim vFolder, vFile, vData, vLine, vAudit, vAbsFile, vFileExtension, vName, vHeader
  Dim oFs, oZip, oFileIn, oFileOut, oStream

  If Request("vOk") = "y" Then sGenerate
    
  Sub sGenerate()

    Server.ScriptTimeout = 60 * 10  '...allow 10 minutes for scripts
  
    vAudit = False
    vAudit = True

    '...bcp does not generate headers, so need to put these at the top of the list
    vName = svMembNo 
    vHeader = "" _
            & "Group" & vbTab _
            & "First Name" & vbTab _
            & "Last Name" & vbTab _
            & "Password" & vbTab _
            & "Expires" & vbTab _
            & "Program ID" & vbTab _
            & "Program Title" & vbTab _
            & "Module ID" & vbTab _
            & "Module Title" & vbTab _
            & "Time Spent" & vbTab _
            & "Bookmark" & vbTab _
            & "Best Date" & vbTab _
            & "Best Score" & vbTab _
            & "No Attempts" & vbTab _
            & "Last Accessed" & vbTab _
            & "Completed"

    '...find the local folder (path)
    vFolder = Server.MapPath(Request.ServerVariables("PATH_INFO"))
    vFolder = Left(vFolder, InstrRev(vFolder, "\"))
    If vAudit Then Response.Write "<p>Local Folder is: " & vFolder

    '...kill all previous files for this account
    vFile = vFolder & vName & ".*"
    If vAudit Then Response.Write "<p>Deleting previous files (" & vFile & ")..."
    Set oFs = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    oFs.DeleteFile(vFile)
    On Error GoTo 0

    vFile   = vFolder & vName & ".tmp" '...this is for the sql data before the header is added

    If vAudit Then Response.Write "<br>Generated text file (length of SQL, which must be less than 1024) is : " & Len(vSql)
    vSql = Replace(Request("History_Sql"), "'", "''")  '...get from cookie and double quote single quotes
    vSql = "MASTER..xp_cmdshell 'bcp """ & vSql & """ queryout " & vFile & " -c -t\t -T -S'"
    If vAudit Then Response.Write "<p>Generating new txt file (" & vSql & ")..."
    sOpenDb
    oDb.Execute vSql
    sCloseDb

    '...wait until we know the .txt file has been created
    vFile = vFolder & vName & ".tmp"
    For j = 1 To 10000
	    If (oFs.FileExists(vFile)) Then Exit For
    Next
    If j > 9999 Then If vAudit Then Response.Write "<p>Temp File was NOT created..." : Exit Sub

    '...open the input sql data file
    Set oFileIn  = oFs.OpenTextFile(vFile, ForReading, False)

    '...open the output data file
    vFile = vFolder & vName & ".txt"
    Set oFileOut = oFs.OpenTextFile(vFile, ForWriting, True)

    '...first insert the headers
    oFileOut.WriteLine vHeader'...put the header in first

    '...add in all the sql data
    Do While oFileIn.AtEndOfStream <> True
      vLine = oFileIn.ReadLine
      If Len(vLine) > 10 Then '...for some reason blank lines seem to appear
        vLine = Replace(vLine, "~", Chr(34))
        oFileOut.WriteLine vLine
      End If
    Loop
    oFileIn.Close

    '...wait until we know the .txt file has been created
    For j = 1 To 10000
	    If (oFs.FileExists(vFile)) Then Exit For
    Next
    If j > 9999 Then If vAudit Then Response.Write "<p>CSV was NOT created..." : Exit Sub


    '...now zip the file for download
    Set oZip                   = CreateObject("PolarZipLight.ZipLight")
    oZip.RecurseSubDirectories = True
    oZip.SourceDirectory       = vFolder
    oZip.FilesToProcess        = vName & ".txt"
    oZip.FilesToExclude        = "*.zip | _*.*"
    oZip.ZipFileName           = vFolder & vName & ".zip"
    oZip.AllowErrorReporting   = True
    oZip.Add


    '...wait until we know the .zip file has been created
    vFile = vFolder & vName & ".zip"
    For j = 1 To 10000
	    If (oFs.FileExists(vFile)) Then Exit For
    Next
    If j > 9999 Then If vAudit Then Response.Write "<p>ZIP was NOT created..." : Exit Sub


    '...stream out the zip
		Set oFileOut = oFs.GetFile(vFile)
		'... first clear the response, and then set the appropriate headers
		Response.Clear
		'... the filename you give it will be the one that is shown
		' to the users by default when they save
		Response.AddHeader "Content-Disposition", "attachment; filename=" & oFileOut.Name
		Response.AddHeader "Content-Length", oFileOut.Size
		Response.ContentType = "application/octet-stream"
		Set oStream = Server.CreateObject("ADODB.Stream")
		oStream.Open
		'... set as binary
		oStream.Type = 1
		Response.CharSet = "UTF-8"
		'... load into the stream the file
		oStream.LoadFromFile(vFile)
		'... send the stream in the response
		Response.BinaryWrite(oStream.Read)
		oStream.Close

    Set oZip      = Nothing  
    Set oFs 			= Nothing
		Set oStream 	= Nothing
 		Set oFileOut 	= Nothing
 		
  End Sub 		

%>
<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="http://vubiz.com/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/jQuery.js"></script>
  <script>
    function sql() {
      location.href="Default.asp?vOk=y&vSql=" + $.cookie("History_Sql");
    };
  </script>    
  <title></title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <div align="center">
    <table border="1" width="60%" cellspacing="0" cellpadding="10" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td align="center">
          <h1 align="center"><br>Download Learners' History</h1>
          <p align="left">This service allows you to download all selected Learners' History Items into a tab delimited text file. Please click <b>Begin</b> once and be patient.</p>
          <p align="center">
            <input type="button" onclick="location.href='<%=Request("vFrom")%>'" value="Return" name="bReturn" class="button85"><%=f10%>
            <input type="button" onclick="sql()" value="Begin" name="bBegin" class="button85"></p>
          <p align="left">Soon after you click <b>Begin</b>, you will be asked to <b>Open</b> or <b>Save</b> your file,&nbsp; Selecting <b>Save</b> will put a zipped <b>TXT</b> file onto your desktop (or where ever you choose to save it).&nbsp; Unzip the file.&nbsp; The unzipped .txt file can then be imported into your local system.<br><br>Note: if you wish to view the file in Excel then right click and select &quot;Open with...&quot; - select Excel).&nbsp; Once open you can save it locally as either &quot;.xls&quot; (Excel 2003) or .&quot;xlsx&quot;&nbsp;(Excel 2007).&nbsp; You may wish to click on the top left of all cells to highlight the entire sheet.&nbsp; Then format the sheet using the &quot;AutoFit Column Width&quot; option.&nbsp; Also it will look more organized if you left justify all columns.</p>
        </td>
      </tr>
    </table>
  </div>
  <p></p>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>