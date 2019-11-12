<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script language="JavaScript" src="/V5/Inc/Functions.js"></script>
  <title>Upload</title>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% 
    Server.Execute vShellHi 
    Dim oUp, oFs, vFileName, vFileNameOk, vSize, vCustId, vInactivate, vAddProgs, vUseGroups

    vCustId      = fDefault(Request("vCustId"), svCustId)
    sGetCust vCustId
    vFileNameOk = Ucase(vCustId & "_Learners.txt")
    vInactivate = fDefault(Request("vInactivate"), "n")
    vAddProgs   = fDefault(Request("vAddProgs"), "n")
    vUseGroups  = fIf(fCritOk(vCust_AcctId), "y", "n")

    Set oUp = Server.CreateObject("SoftArtisans.FileUp")
    Set oFs = Server.CreateObject("Scripting.FileSystemObject")   

    If oUp.Form.Count > 0 Then

      vCustId     = oUp.Form("vCustId")
      vInactivate = fDefault(oUp.Form("vInactivate"), "n")
      vAddProgs   = fDefault(oUp.Form("vAddProgs"), "n")
      sGetCust vCustId

      vFileName = oUp.UserFilename
      If Len(vFileName) = 0 Then
        Response.Redirect "/V5/Code/Error.asp?vCustId=" & vCustId & "&vInactivate=" & vInactivate & "&vAddProgs=" & vAddProgs & "&vUseGroups=" & vUseGroups & "&vReturn=/V5/Repository/Upload/Upload_Advanced.asp&vErr=" & Replace("No file has been selected.", " ", "+")
      End If
      vFileName = Ucase(Mid(vFileName, InstrRev(vFileName, "\") + 1))
      If vFileName <> vFileNameOk Then
        Response.Redirect "/V5/Code/Error.asp?vCustId=" & vCustId & "&vInactivate=" & vInactivate & "&vAddProgs=" & vAddProgs & "&vUseGroups=" & vUseGroups & "&vReturn=/V5/Repository/Upload/Upload_Advanced.asp&vErr=" & Replace("Please browse to find the file named '" & vFileNameOk & "'.", " ", "+")
      End If
      vSize = oUp.TotalBytes

      '...delete file if it exists (Server 2008/IIS7 issue)
      If oFs.FileExists(Server.MapPath(vFileName)) Then
        oUp.Delete Server.MapPath(vFileName)
      End If

      oUp.SaveInVirtual vFileName
      Set oUp = Nothing
      If Err = 0 Then 
        Response.Redirect "Upload_Advanced_Ok.asp?vCustId=" & vCustId & "&vInactivate=" & vInactivate & "&vAddProgs=" & vAddProgs 
      Else
        Response.Redirect "/V5/Code/Error.asp?vCustId=" & vCustId & "&vInactivate=" & vInactivate & "&vAddProgs=" & vAddProgs & "&vUseGroups=" & vUseGroups & "&vReturn=/V5/Repository/Upload/Upload_Advanced.asp&vErr=" & Replace("Your file could not be uploaded because:<br><br>" & Err.Description & ".", " ", "+")
      End If
    End If

  %>
  <div align="center">
  <table border="0" cellpadding="10" style="border-collapse: collapse" bordercolor="#DDEEF9" width="80%">
    <form enctype="multipart/form-data" method="post" action="Upload_Advanced.asp">
      <tr>
        <td align="center">
        <h1>Upload Learner Profiles (Advanced)</h1>
        <a class="c2" href="#" onclick="toggle('divDetails')">Click here for details</a>.<br><br>

        <div id="divDetails" class="div" align="left">
          <br>You can upload any number of Learner profiles from a Tab Delimited Text File stored on your system as <font color="#FF0000"><%=vCust_Id%>_LEARNERS.txt</font> (not case sensitive).&nbsp; (Note: Do NOT upload a comma separated values file). New learner profiles will be added to the account while existing learners will have their profile updated accordingly. All profiles uploaded are deemed to be &quot;Learners&quot;, i.e. you cannot upload Facilitators, Managers, etc.<br><br><b>First</b>: Start with an Excel spreadsheet of your learners and arrange the columns so they conform to the sample file here: <a target="_blank" class="c2" href="DEMO0000_LEARNERS.xls"><font color="#3B5E91">DEMO0000_LEARNERS.xls</font></a>. Do not remove any unused columns from this spreadsheet. Leave a column blank if it is not applicable, but do not remove it.&nbsp; Your spreadsheet MUST contain a header row that MUST be exactly the same as in the sample file. If you use Programs or Jobs, ensure you separate multiple entries with a space.<br><br><b>Second</b>: Save the .xls file as a Text (Tab Delimited) file for uploading. Click here for a sample text file: <a target="_blank" href="DEMO0000_LEARNERS.txt"><font color="#3B5E91">DEMO0000_LEARNERS.txt</font></a>. <br><br><b>Third</b>: Follow the steps below to upload your .txt file. (Note: Step 2 is only a choice for Custom / Corporate. If you are updating your Learner database by uploading ALL active Learners, you should “Inactivate all existing learners” first, leaving just the Learners being uploaded as active.<p>These are the fields in the file upload and must appear in this order:&nbsp; <br>&nbsp;
          </p>
          <div align="right">
            <table border="1" width="90%" cellspacing="0" cellpadding="4" style="border-collapse: collapse" bordercolor="#DDEEF9">
              <tr>
                <th align="left" class="c1" valign="top">Column</th>
                <th align="left" class="c1" valign="top">Field</th>
                <th align="left" class="c1" valign="top">&nbsp;</th>
                <td align="left" class="c1" valign="top">Description</td>
                <td align="left" class="c1" valign="top" nowrap>Example</td>
              </tr>
              <tr>
                <th align="left" class="c1" valign="top">1</th>
                <th align="left" class="c1" valign="top" nowrap>Learner ID</th>
                <td align="left" valign="top">Mandatory</td>
                <td valign="top">Can only contain characters <i>A-Z</i>, <i>0-9</i> and <i>-_@.</i> (ie email rules). Note for Self-Serve sites, this field is called &quot;Password&quot;.&nbsp; </td>
                <td valign="top" nowrap>A_12345</td>
              </tr>
              <tr>
                <th align="left" class="c1" valign="top">2</th>
                <th align="left" class="c1" valign="top" nowrap>Group ID</th>
                <td align="left" valign="top"><%=fIf(vUseGroups, "Mandatory", "Not used")%></td>
                <td valign="top">Only used in Custom (Corporate) sites. Not used in Self-Serve.</td>
                <td valign="top" nowrap><%=fIf(vUseGroups, "HO", "")%></td>
              </tr>
              <tr>
                <th align="left" class="c1" valign="top">3</th>
                <th align="left" class="c1" valign="top" nowrap>First Name</th>
                <td align="left" valign="top">Mandatory</td>
                <td valign="top">&nbsp;</td>
                <td valign="top" nowrap>Jean</td>
              </tr>
              <tr>
                <th align="left" class="c1" valign="top">4</th>
                <th align="left" class="c1" valign="top" nowrap>Last Name</th>
                <td align="left" valign="top">Mandatory</td>
                <td valign="top">&nbsp;</td>
                <td valign="top" nowrap>Smith</td>
              </tr>
              <tr>
                <th align="left" class="c1" valign="top">5</th>
                <th align="left" class="c1" valign="top" nowrap>Email Address:</th>
                <td align="left" valign="top">Mandatory if issuing email alerts; Optional otherwise</td>
                <td valign="top">This field is not compared to existing values and is mandatory for those issuing email alerts.</td>
                <td valign="top" nowrap>jean.smith@email.com</td>
              </tr>
              <tr>
                <th align="left" class="c1" valign="top">6</th>
                <th align="left" class="c1" valign="top" nowrap>Password</th>
                <td align="left" valign="top"><%=fIf(vCust_Pwd, "Mandatory", "Not used")%></td>
                <td valign="top">Optional in Custom (Corporate) sites. Not used in Self-Serve.</td>
                <td valign="top" nowrap><%=fIf(vCust_Pwd, "Smiley.Face", "")%></td>
              </tr>
              <tr>
                <th align="left" class="c1" valign="top">7</th>
                <th align="left" class="c1" valign="top" nowrap>Programs</th>
                <td align="left" valign="top">Optional</td>
                <td valign="top">If used, enter Program IDs separated by <b>spaces</b>.</td>
                <td valign="top" nowrap>P1234EN P1235EN P9995EN</td>
              </tr>
              <tr>
                <th align="left" class="c1" valign="top">8</th>
                <th align="left" class="c1" valign="top" nowrap>Memo</th>
                <td align="left" valign="top">Optional</td>
                <td valign="top">If used and if values contain more than one field, separate by pipes</td>
                <td valign="top" nowrap>Scranton|PA</td>
              </tr>
              <tr>
                <th align="left" class="c1" valign="top">9</th>
                <th align="left" class="c1" valign="top" nowrap>Jobs</th>
                <td align="left" valign="top">Optional</td>
                <td valign="top">If used, enter Job IDs | Program IDs separated by <b>comma space</b>.</td>
                <td valign="top" nowrap>J0011EN|P1218EN, J0012EN|P2256EN</td>
              </tr>
              <tr>
                <th align="left" class="c1" valign="top">10</th>
                <th align="left" class="c1" valign="top" nowrap>Jobs</th>
                <td align="left" valign="top">Optional</td>
                <td valign="top">If used, enter Job IDs | Program IDs separated by <b>comma space</b>.</td>
                <td valign="top" nowrap>J0011EN|P1218EN, J0012EN|P2256EN</td>
              </tr>
            </table>
            <br>
          </div>
        </div>


        <table border="1" cellpadding="5" style="border-collapse: collapse" bordercolor="#DDEEF9">
          <tr>
            <td valign="top">
              <b><font size="4">1.</font></b></td>
            <td valign="top">
              Click <b>Browse</b> to find your <%=vCust_Id%>_LEARNERS.txt file on your system...<br><br> 
              <input type="file" name="vEmployees" size="35" class="button">
            </td>
          </tr>
          <tr>
            <td valign="top">
              <b><font size="4">2.</font></b></td>
            <td valign="top">
              If you always upload <b>all</b> active learners each time, then...<br><br>
              <input type="checkbox" name="vInactivate" value="y" <%=fcheck("y", vinactivate)%>>Inactive all existing Learners before uploading the (active) learners.
            </td>
          </tr>
          <tr>
            <td valign="top">
              <font size="4">
              <b>3.</b></font></td>
            <td valign="top">
              If you upload Program Ids, then specify...<br><br>
              <input type="checkbox" name="vAddProgs" value="y" <%=fcheck("y", vAddProgs)%>>Add uploaded Program IDs to any that might be on file. <br><br><font color="#FF0000">WARNING:</font> 
              <ul>
                <li>If you do not check this box, Program IDs already on file for a learner will be replaced instead of added. </li>
                <li>Ensure you do not upload an existing Program ID already assigned to a Learner’s profile. You will NOT get an error message if you do this, so please check your Programs Purchased and Assigned Report carefully before upload to make sure you are not duplicating Program IDs for a Learner. On a Self-Serve site, uploading a duplicate Program ID will not deplete your inventory of programs available, but it will reassign the program to the Learner’s profile and trigger an email alert (if the feature is enabled on a Self-Serve account).</li>
                <li>This option does NOT apply to Job Ids which are ALWAYS replaced.</li>
              </ul>
            </td>
          </tr>
          <tr>
            <td valign="top">
              <b><font size="4">4.</font></b></td>
            <td valign="top">
              Click <b>Submit</b> to upload the file.<b><font color="#FF0000"> NOTE: ONLY CLICK ONCE!</font></b><br><br>
              <b><font color="#FF0000">
              <input type="submit" value="Submit" class="button100">
            </font></b>
            </td>
          </tr>
          </table>
        <h6 align="center">Note: This service can take up to 30 minutes if there are more than 1000 records.</h6>
        </td>
      </tr>
      <input type="hidden" name="vCustId" value="<%=vCustId%>">
      <tr>
        <td align="center">&nbsp;</td>
      </tr>
    </form>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
