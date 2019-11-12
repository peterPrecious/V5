<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<html>

<head>
  <title>Download/Default</title>
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
</head>

<body>

  <!--#include virtual = "V5/Inc/Shell_HiSolo.asp"-->

  <% 
			'... service requires any 8 char id (typicall cust id like VUBZ2274 - it bypasses security
			Dim oUp, vFileName, vAccessId
	    Set oUp = Server.CreateObject("SoftArtisans.FileUp")
	    If oUp.Form.Count > 0 Then  
	      Server.ScriptTimeout = 60 * 60         '...allow 60 minutes
	      oUp.MaxBytes         = 0               '...no limit on individual file size
	      vAccessId = Ucase(oUp.Form("vAccessId"))
		    If Len(vAccessId) <> 8 Then
  	      Response.Write "<div style='text-align:center; padding:20px;'><h5>That is not a valid Access Id!</h5><br><input onclick='window.history.back()' type='button' value='Return' class='button'></div>"
		    Else		      
		      vFileName = oUp.UserFilename
		      vFileName = vAccessId & "_" & Mid(vFileName, InstrRev(vFileName, "\") + 1)	
		      On Error Resume Next 
		      oUp.SaveInVirtual vFileName 
		      If Err = 0 Then 
		        Response.Write "<div style='text-align:center; padding:20px;'><h2>Thank you. Your file was uploaded successfully.</h2></ br><input onclick='window.history.back()' type='button' value='Return' class='button'></div>"
		      Else
		        Response.Write "<div style='text-align:center; padding:20px;'><h5>Your file could not be uploaded because:<br><br>" & Err.Description & ".</h5></ br><input onclick='window.history.back()' type='button' value='Return' class='button'></div>"
		      End If
		      On Error Goto 0      
		      Set oUp = Nothing
				End If
	    Else
  %>
  <div style="text-align: center">
    <img border="0" src="/v5/images/Logos/vubz.jpg" height="50">
  </div>
  <h1>VUBIZ Generic File Upload</h1>

  <div style="width: 500px; margin: auto;">
    <p class="c2">Notes:</p>
    <ol>
      <li>This service enables you to upload any file to the Vubiz servers.</li>
      <li>For security reasons, access this service via <a href="https://learn.vubiz.com">https://learn.vubiz.com</a>.</li>
      <li>If the file you upload already exists, it will be overwritten.</li>
      <li>To begin, enter your Access Id then click <b>Browse</b> to find your local file...</li>
      <li>Then click <b>Submit</b> to upload it.</li>
    </ol>
  </div>

  <form enctype="multipart/form-data" method="post" action="Default.asp">
    <table class="table" style="text-align: center">
      <tr>
        <th style="width: 50%">Access Id :  </th>
        <td>
          <input type="text" name="vAccessId" size="20" xvalue="passé"></td>
      </tr>
      <tr>
        <th>File : </th>
        <td>
          <input type="file" name="vFile"></td>
      </tr>
    </table>

    <div style="margin: 20px; text-align: center;">
      <input type="submit" value="Submit" class="button">
    </div>
  </form>

  <% 
		    End If 
  %>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
