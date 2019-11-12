<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  Const ForReading = 1, ForWriting = 2
  Dim vFolder, vFile, vData, vLine, vAudit, vAbsFile, vFileExtension, vName, vStrDate, vEndDate, vScript
  Dim oFs, oZip, oFileIn, oFileOut, oStream

  
  If Request("vOk") = "y" Then 
    vStrDate = Request("vStrDate")
    vEndDate = Request("vEndDate")
    If DateDiff ("d", vStrDate, vEndDate) > 31 Then
      vScript = "onload=""alert('The date range cannot be greater than 1 month.')"""
    ElseIf DateDiff ("d", vStrDate, vEndDate) < 1 Then
      vScript = "onload=""alert('The start date is beyond the end date.')"""
    Else
      vEndDate = fFormatSqlDate(DateAdd("d", 1, vEndDate)) '...add on another day to ensure we get all of the last day.
      vScript = ""
      vName = svCustId & "." & vStrDate & "-" & vEndDate
      vName = Replace(vName, " ", ".") 
      vName = Replace(vName, ",", "") 
      sGenerate
    End If
  End If


  Sub sGenerate()
    vAudit = False

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

    vFile   = vFolder & vName & ".txt" '...this is the raw data before the header is added

    vSql = "" _
         & " SELECT TOP 100 PERCENT" _
         & "    [Group 1],					"_
         & "    [Group 2], 					"_
         & "    [Password], 				"_
         & "    [Last Name], 				"_
         & "    [First Name], 			"_
         & "    [Program ID], 			"_
         & "    [Program Title],		"_ 
         & "    [Assessment ID], 		"_
         & "    [Assessment Title],	"_ 
         & "    [Last Score], 			"_
         & "    [Best Score], 			"_
         & "    [No Attempts],			"_
         & "    [Time Spent], 			"_
         & "    REPLACE([Memo], ''|'', '','') "_
         & " FROM V5_Vubz.dbo.vLRC "_
         & " WHERE [Acct Id]= ''" & svCustAcctId & "''" _
         & " AND [Last Score] BETWEEN CAST(''" & vStrDate & "'' AS DATETIME) AND CAST(''" & vEndDate & "'' AS DATETIME) "


    If vAudit Then Response.Write "<br>Generated text file (length of SQL, which must be less than 1024) is : " & Len(vSql)
    vSql = "MASTER..xp_cmdshell 'bcp """ & vSql & """ queryout " & vFile & " -c -t, -T -S'"
    If vAudit Then Response.Write "<p>Generating new txt file (" & vSql & ")..."
    sOpenDb
    oDb.Execute vSql
    sCloseDb

    '...wait until we know the .txt file has been created
    vFile = vFolder & vName & ".txt"
    For j = 1 To 10000
	    If (oFs.FileExists(vFile)) Then Exit For
    Next
    If j > 9999 Then If vAudit Then Response.Write "<p>CSV was NOT created..." : Exit Sub

    '...grab the header file to merge into the final file
    vFile = vFolder & "Headers.txt"
    If vAudit Then Response.Write "<p>Reading in the header file to prep for copy (" &  vFile & ")..."
    Set oFileIn = oFs.OpenTextFile(vFile, ForReading)
    vLine = oFileIn.ReadLine
    If vAudit Then Response.Write "<p>Reading in the data file to prep for copy (" &  vFile & ") ..."
    If vAudit Then Response.Write "<p>...and writing out the merged file (" &  vFile & ") ..."
    vFile = vFolder & vName & ".txt"
    Set oFileIn  = oFs.OpenTextFile(vFile, ForReading, False)
    vFile = vFolder & vName & ".csv"
    Set oFileOut = oFs.OpenTextFile(vFile, ForWriting, True)
    oFileOut.WriteLine vLine '...put the header in first
    '...this is where we replace tildes with commas - not used in this version
    Do While oFileIn.AtEndOfStream <> True
      vLine = oFileIn.ReadLine
      If Len(vLine) > 10 Then '...for some reason blank lines seem to appear
        oFileOut.WriteLine vLine
      End If
    Loop
    oFileIn.Close


    '...wait until we know the .csv file has been created
    vFile = vFolder & vName & ".csv"
    For j = 1 To 10000
	    If (oFs.FileExists(vFile)) Then Exit For
    Next
    If j > 9999 Then If vAudit Then Response.Write "<p>CSV was NOT created..." : Exit Sub


    '...now zip the file for download
    Set oZip                   = CreateObject("PolarZipLight.ZipLight")
    oZip.RecurseSubDirectories = True
    oZip.SourceDirectory       = vFolder
    oZip.FilesToProcess        = vName & ".csv"
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
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <script language="JavaScript" src="/V5/Inc/Launch.js"></script>
  <script language="JavaScript" src="/V5/Inc/Calendar.js"></script>
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <title></title>
</head>

<body <%=vScript%> leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <div align="center">
    <table border="1" width="80%" cellspacing="0" cellpadding="2" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td>
          <h1 align="center"><br>Download Learners and Related Learning Activities</h1>
          <p align="left">This service allows you to download all Learners and their Learning Activities in CSV format that occurred during any month (or less).&nbsp; It can take several minutes.&nbsp; Please only click <b>Begin</b> once and be patient.</p>
          <form method="POST" action="Default.asp" name="fForm" id="fForm">
            <table border="0" width="100%" cellspacing="0" cellpadding="3">
              <%
                '...get previous month's start and end date
                If vStrDate = "" Or Not IsDate(vStrDate) Then 
                  vStrDate = fFormatSqlDate(DateAdd("m", -1, Now))
                  vStrDate = MonthName(Month(vStrDate), True) & " 1, " & Year(vStrDate)
                End If
                If vEndDate = "" Or Not IsDate(vEndDate)  Then 
                  vEndDate = fFormatSqlDate(DateAdd("d", -1, (DateAdd("m", 1, vStrDate))))
                End If              
              %>           
              <tr>
                <th align="right" width="50%">Start Date : </th>
                <td width="50%">
                  <input type="text" name="vStrDate" id="vStrDate" size="12" value="<%=vStrDate%>" style="text-align: center" class="c2">&nbsp;
                  <a href="javascript:show_calendar('vStrDate','EN', '<%=Month(vStrDate)-1%>', '<%=Year(vStrDate)%>');"><img border="0" src="/V5/Images/Icons/Calendar.jpg" align="absbottom"></a>
                </td>
              </tr>
              <tr>
                <th align="right" width="50%">End Date : </th>
                <td width="50%">
                  <input type="text" name="vEndDate" id="vEndDate" size="12" value="<%=vEnddate%>" style="text-align: center" class="c2">&nbsp;
                  <a href="javascript:show_calendar('vEndDate','EN', '<%=Month(vEndDate)-1%>', '<%=Year(vEndDate)%>');"><img border="0" src="/V5/Images/Icons/Calendar.jpg" align="absbottom"></a>
                </td>
              </tr>
              <tr>
                <td colspan="2" align="center">
                <p><br><br><input type="submit" value="Begin" name="bBegin" class="button"><br>&nbsp;</p></td>
              </tr>
            </table>
            <input type="hidden" name="vOk" value="y">
          </form>
          <p align="center" class="d2"><a href="#" onclick="toggle('divHelp')"><font color="#008000">Click here for details on this service.</font></a><font color="#008000"><br>&nbsp;</font></p>
          <div class="div" id="divHelp">
          <p align="left">After you enter the date range, click <b>Begin</b> where, after a few minutes, you will be asked to <b>Open</b> or <b>Save</b> your file,&nbsp; Select <b>Save</b>.&nbsp;This will put a zipped <b>CSV</b> file onto your desktop (or where ever you choose to save it).&nbsp; Open the file in Excel (this should happen when you double click on the file).&nbsp; After using this file, save it as a local &quot;.xls&quot; or .&quot;xlsx&quot;&nbsp;file.</p>
          <p class="c6">If the number of records generated exceed 65,000 you will need Excel 2007.</p>
          <p>When in Excel, click on the top left of all cells to highlight the entire sheet.&nbsp; Then format the sheet using the &quot;AutoFit Column Width&quot; option.&nbsp; Also it may look more organized if you left justify all columns.<br><br>If you use the &quot;Memo&quot; field for organizational purposes, this field will appear at the far right column of the sheet.&nbsp; It may contain cells that have data separated by pipes (&quot;|&quot;).&nbsp; To put these values into their own columns, highlight the Memo column then choose the Data function &quot;Text to Columns with the pipe delimiter.&nbsp; For example, if one memo field contained &quot;Montreal|Canada&quot; you&#39;d ask the wizard to break this up as follows:</p>
          <p align="center"><img border="0" src="../csv/Suggestion.jpg"><br>&nbsp;</p>
          </div>
        </td>
      </tr>
    </table>
  </div>
  <p></p>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>