<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->

<html>

<head>
  <title>Upload3</title>
  <meta charset="UTF-8">
  <script src="http://code.jquery.com/jquery-latest.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% 
    Server.Execute vShellHi 
    Dim oUp, oFs, vFileName, vFileNameOk, vSize, vCustId, vAction, vDaysOk, vAddProgs, vUseGroups, vReportsTo

    vCustId      		= fDefault(Request("vCustId"), svCustId)
    sGetCust vCustId
    vFileNameOk 		= Ucase(vCustId & "_Learners.txt")
    vAction     		= fDefault(Request("vAction"), "x")
    vDaysOk         = fDefault(Request("vDaysOk"), 0)
    vAddProgs   		= fDefault(Request("vAddProgs"), "n")
    vUseGroups 			= fIf(fCritOk(vCust_AcctId), "y", "n")
    vReportsTo      = fIf(vCust_ChannelReportsTo, "y", "n")
    
    Set oUp = Server.CreateObject("SoftArtisans.FileUp")
    Set oFs = Server.CreateObject("Scripting.FileSystemObject")   

    If oUp.Form.Count > 0 Then

      vCustId     			= oUp.Form("vCustId")
      vAction     		  = fDefault(oUp.Form("vAction"), "")
	    vDaysOk           = fDefault(oUp.Form("vDaysOk"), 0)
      vAddProgs   			= fDefault(oUp.Form("vAddProgs"), "n")
      vReportsTo        = fDefault(oUp.Form("vReportsTo"), "n")
      sGetCust vCustId

      vFileName = oUp.UserFilename
      If Len(vFileName) = 0 Then
        Response.Redirect "/V5/Code/Error.asp?vCustId=" & vCustId & "&vAction=" & vAction & "&vDaysOk=" & vDaysOk & "&vAddProgs=" & vAddProgs & "&vUseGroups=" & vUseGroups & "&vReturn=/V5/Repository/Upload3/Upload3.asp&vErr=" & Replace("No file has been selected.", " ", "+")
      End If
      vFileName = Ucase(Mid(vFileName, InstrRev(vFileName, "\") + 1))
      If vFileName <> vFileNameOk Then
        Response.Redirect "/V5/Code/Error.asp?vCustId=" & vCustId & "&vAction=" & vAction & "&vDaysOk=" & vDaysOk & "&vAddProgs=" & vAddProgs & "&vUseGroups=" & vUseGroups & "&vReturn=/V5/Repository/Upload3/Upload3.asp&vErr=" & Replace("Please browse to find the file named '" & vFileNameOk & "'.", " ", "+")
      End If
      vSize = oUp.TotalBytes

      '...delete file if it exists (Server 2008/IIS7 issue)
      If oFs.FileExists(Server.MapPath(vFileName)) Then
        oUp.Delete Server.MapPath(vFileName)
      End If

      oUp.SaveInVirtual vFileName
      Set oUp = Nothing
      If Err = 0 Then 
        Response.Redirect "Upload3_Ok.asp?vCustId=" & vCustId & "&vAction=" & vAction &  "&vDaysOk=" & vDaysOk & "&vAddProgs=" & vAddProgs & "&bRportsTo=" & vReportsTo
      Else
        Response.Redirect "/V5/Code/Error.asp?vCustId=" & vCustId & "&vAction=" & vAction &  "&vDaysOk=" & vDaysOk & "&vAddProgs=" & vAddProgs & "&vUseGroups=" & vUseGroups & "&vReturn=/V5/Repository/Upload3/Upload3.asp&vErr=" & Replace("Your file could not be uploaded because:<br><br>" & Err.Description & ".", " ", "+")
      End If
    End If

  %>
  <table style="width: 80%; margin: auto;">
    <form id="fUpload" enctype="multipart/form-data" method="post" action="Upload3.asp">
      <tr>
        <td style="text-align: center">
          <h1>Upload Learner Profiles (Advanced with ReportsTo)</h1>
          <a class="c2" href="#" onclick="toggle('divDetails')">Click for details</a>.<br>
          <br>
          <div id="divDetails" class="div" style="text-align: left;">
            <br>You can upload any number of Learner profiles from a <b>Tab Delimited Text File</b> stored on your system as <b><%=vCust_Id%>_LEARNERS.txt</b> (not case sensitive).&nbsp; <b>Do not upload a comma separated values (CSV) file</b>. New learner profiles will be added to the account while existing learners will have their profile updated accordingly. All profiles uploaded are deemed to be &quot;Learners&quot;, i.e. you cannot upload Facilitators, Managers, etc.<br><br>If you wish, you can <a href="Download3.asp">click here</a> and the system will extract <b>ALL Active Learners</b> from your site which you can save on your desktop as <b><%=vCust_Id%>_LEARNERS.xls</b> and use that to create a tab delimited file for uploading.<br><br>
            <span class="c2">First</span>: Start with an Excel spreadsheet of your learners and arrange the columns so they conform to the sample file here: <a target="_blank" href="DEMO0000_LEARNERS.xls"><font color="navy">DEMO0000_LEARNERS.xls</font></a>. Do not remove any unused columns from this spreadsheet. Leave a column blank if it is not applicable, but do not remove it.&nbsp; Your spreadsheet MUST contain a header row that MUST be exactly the same as in the sample file. If you use Programs or Jobs, ensure you separate multiple entries with a space.<br>
            <br>
            <span class="c2">Second</span>: Save the .xls file as a Text (Tab Delimited) file for uploading. Click here for a sample text file: <a target="_blank" href="DEMO0000_LEARNERS.txt">DEMO0000_LEARNERS.txt</a>. <br>
            <br>
            <span class="c2">Third</span>: Follow the steps below to upload your .txt file. (Note: Step 2 is only a choice for Custom / Corporate. If you are updating your Learner database by uploading ALL active Learners, you should "Inactivate all existing learners" first, leaving just the Learners being uploaded as active.<br /><br />These are the fields in the file upload and must appear in this order:<br /><br /><br />
            <div style="text-align: right">
              <table>
                <tr>
                  <td class="c2">Col</td>
                  <td class="c2" style="width:100px;">Name</td>
                  <td class="c2"></td>
                  <td class="c2">Description</td>
                  <td class="c2">Example</td>
                </tr>
                <tr>
                  <td style="text-align:center" class="c2">1</td>
                  <td class="c2">Learner ID</td>
                  <td>Mandatory</td>
                  <td>Can only contain characters <i>A-Z</i>, <i>0-9</i> and <i>-_@.</i> (ie email rules). Note for Self-Serve sites, this field is called &quot;Password&quot;.&nbsp; </td>
                  <td>A_12345</td>
                </tr>
                <tr>
                  <td class="c2">2</td>
                  <td class="c2">Group ID</td>
                  <td><%=fIf(vUseGroups, "Mandatory", "Not used")%></td>
                  <td>Only used in Custom (Corporate) sites. Not used in Self-Serve.</td>
                  <td><%=fIf(vUseGroups, "HO", "")%></td>
                </tr>
                <tr>
                  <td class="c2">3</td>
                  <td class="c2">First Name</td>
                  <td>Mandatory</td>
                  <td>&nbsp;</td>
                  <td>Jean</td>
                </tr>
                <tr>
                  <td class="c2">4</td>
                  <td class="c2">Last Name</td>
                  <td>Mandatory</td>
                  <td>&nbsp;</td>
                  <td>Smith</td>
                </tr>
                <tr>
                  <td style="text-align:center" class="c2">5</td>
                  <td class="c2">Email Address</td>
                  <td>Mandatory if issuing email alerts; Optional otherwise</td>
                  <td>This field is not compared to existing values and is mandatory for those issuing email alerts.</td>
                  <td>jean.smith@email.com</td>
                </tr>
                <tr>
                  <td style="text-align:center" class="c2">6</td>
                  <td class="c2">Password</td>
                  <td><%=fIf(vCust_Pwd, "Mandatory", "Not available")%></td>
                  <td>Required if site configured to use Passwords (Custom sites).&nbsp; Only used for NEW learners NOT when updating existing learners.</td>
                  <td><%=fIf(vCust_Pwd, "Smiley.Face", "")%></td>
                </tr>
                <tr>
                  <td style="text-align:center" class="c2">7</td>
                  <td class="c2">Programs</td>
                  <td>Optional</td>
                  <td>If used, enter Program IDs separated by <b>spaces</b>.</td>
                  <td>P1234EN P1235EN P9995EN</td>
                </tr>
                <tr>
                  <td style="text-align:center" class="c2">8</td>
                  <td class="c2">Memo</td>
                  <td>Optional</td>
                  <td>If used and if values contain more than one field, separate by pipes</td>
                  <td>Scranton|PA</td>
                </tr>
                <tr>
                  <td style="text-align:center" class="c2">9</td>
                  <td class="c2">Jobs</td>
                  <td>Optional</td>
                  <td>If used, enter Job IDs | Program IDs separated by <b>comma space</b>.</td>
                  <td>J0011EN|P1218EN, J0012EN|P2256EN</td>
                </tr>
                <tr>
                  <td style="text-align:center" class="c2">10</td>
                  <td class="c2">Reports To</td>
                  <td><%=fIf(vReportsTo = "y", "Optional", "Not available")%></td>
                  <td>Available for designated Self-Serve (Channel) accounts.&ensp;If used, enter an EXISTING facilitator learner ID who manages this learner group.</td>
                  <td>Leave empty 
                      <% If vReportsTo = "y" Then %>                     
                      or use one of the Facilitor Ids<br />below [in square brackets]<br />
                    <select size="1" name="vFacs">
                      <%=fMembFacsDropdown (0) %>
                    </select>
                    <% End If %>
                  </td>
                </tr>
              </table>
              <br>
            </div>
          </div>
          <table id="meat">
            <tr>
              <td class="c2">1.</td>
              <td>Click <b>Browse</b> to find your <%=vCust_Id%>_LEARNERS.txt file on your system...<br><br>
                <input type="file" name="vEmployees" size="35" class="button">
              </td>
            </tr>
            <tr>
              <td class="c2">2.</td>
              <td>If you are uploading <b>ALL</b> active learners, then you can select <b>ONE </b>of the following...<br><br>
                <input type="radio" name="vAction" value="x" <%=fcheck("x", vAction)%>>Do not modify the learner file before uploading.<br />
                <input type="radio" name="vAction" value="i" <%=fcheck("i", vAction)%>>Inactive all existing Learners before uploading the active learners. <br />
                <input type="radio" name="vAction" value="s" <%=fcheck("s", vAction)%>>Inactive all existing Learners before uploading,<span class="red">except those that have been added within the last <input type="text" name="vDaysOk" value="<%=vDaysOk%>" style="width: 20px" id="vDaysOk" /> days.</span><br>
                <input type="radio" name="vAction" value="d" <%=fcheck("d", vAction)%>><span class="red">Delete all existing Learners before uploading the active learners (exercise caution)</span>.
              </td>
            </tr>
            <tr>
              <td class="c2">3.</td>
              <td>If you Upload Program Ids, then specify...<br>
                <br>
                <input type="checkbox" name="vAddProgs" value="y" <%=fcheck("y", vaddprogs)%>>Add Uploaded Program IDs to any that might be on file. <br>
                <br>
                <h6>WARNING</h6>
                <ul>
                  <li>If you do not check this box, Program IDs already on file for a learner will be replaced instead of added. </li>
                  <li>Ensure you do not upload an existing Program ID already assigned to a Learner’s profile. You will NOT get an error message if you do this, so please check your Programs Purchased and Assigned Report carefully before upload to make sure you are not duplicating Program IDs for a Learner. On a Self-Serve site, uploading a duplicate Program ID will not deplete your inventory of programs available, but it will reassign the program to the Learner’s profile and trigger an email alert (if the feature is enabled on a Self-Serve account).</li>
                  <li>This option does NOT apply to Job Ids which are ALWAYS replaced.</li>
                </ul>
              </td>
            </tr>

            <script>
                  function jSubmitPlusX (formId, hideId, showId) {
                    var ok = true;
                    if ($("input:radio[name=vAction]:checked").val() == "s") {
                      ok = false;
                      if (isNumber($("#vDaysOk").val())) {
                        if (($("#vDaysOk").val() > 0) & ($("#vDaysOk").val() < 365)) {
                          ok = true;
                        } else {
                          $("#vDaysOk").focus();
                        }
                      }
                    }
  
                    if (ok) {
                      jSubmitPlus(formId, hideId, showId);
                    } else {
                      alert ("No of days must be between 1 and 365");
                      return false;
                    }
                  }
            </script>

            <tr>
              <td class="c2">4.</td>
              <td>Click <b>Submit</b> to upload the file...<br /><br>
                <div style="text-align: center">
                  <input id="bSubmit" onclick="jSubmitPlusX('fUpload', 'bSubmit', 'iLoad')" type="button" value="Submit" class="button100">
                  <img class="div" id="iLoad" src="/V5/Images/Common/ProgressBar.gif" />
                </div>
              </td>
            </tr>

          </table>
          <h6>Note: This service can take several minutes if there are numerous records.</h6>
        </td>
      </tr>
      <input type="hidden" name="vCustId" value="<%=vCustId%>">
    </form>
  </table>

   <style>
   #meat td {padding:10px;}
  </style>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
