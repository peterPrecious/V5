<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<%
  If svMembLevel < 3 Then Response.Redirect "Default.asp" 
%>

<html>

<head>
  <title>DocumentUpload</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table class="table">
    <tr>
      <td colspan="2">
        <h1>Upload Custom Document</h1>
        <p>This allows you to upload a document that is customized to your Organization&#39;s standards.&nbsp; Uploaded documents appear when learners select VuBuild <strong>smartlinks</strong>.&nbsp;&nbsp; If you offer this document in different languages, please develop ONE document per language, ie do not make the document multi-lingual. Note that the file name you use to upload can be any name - it will be converted to the name you select below.&nbsp; If you update any document in the future, simply re-upload the revised document and it will over ride whatever has been previously uploaded.</p>
        <p>Select the Document Name, Language then Next to Upload. <strong><em>&nbsp;You will then be directed to a new page to upload your document.&nbsp; Once uploaded the page will immediately return you back here - signifying that the upload was successful.</em></strong></p>
        <p>&nbsp;</p>
        <form method="POST" action="/docservice/default.aspx">
          <table class="table">
            <tr>
              <th>Document File Name :</th>
              <td>
                <input type="radio" value="harassment.pdf"        name="vDocument" checked="checked">harassment.pdf    <br />
                <input type="radio" value="conflict.pdf"          name="vDocument">conflict.pdf                        <br />
                <input type="radio" value="reaffirmation.pdf"     name="vDocument">reaffirmation.pdf                   <br />
                <input type="radio" value="ethicsemployees.pdf"   name="vDocument">ethicsemployees.pdf                                
              </td>
            </tr>
            <tr>
              <th>For Customer Id : </th>
              <td>
                <% If svMembLevel < 5 Then %> 
                  <%=svCustId%>
                  <input type="hidden" name="vCustId" value="<%=svCustId%>">
                <% Else %>
                  <input type="text" name="vCustId" value="<%=svCustId%>" style="width: 75px"><br>
                    Note: to upload a Master 
                      for ALL accounts leave empty, 
                      for this account, or any other, just enter the 4 character ID (ie <%=Left(svCustId, 4)%>) 
                      otherwise leave as an 8 character ID which limits access to this document to this account. 
                <% End If %> 
              </td>
            </tr>
            <tr>
              <th>Language :</th>
              <td>
                <input type="radio" value="EN" name="vLang" checked="checked">EN
                <input type="radio" value="ES" name="vLang">ES
                <input type="radio" value="FR" name="vLang">FR </td>
            </tr>
            <tr>
              <td colspan="2" style="text-align: center;">
                <input style="margin: 30px;" class="button" type="submit" value="Next" name="bNext">
              </td>
            </tr>
          </table>

<!--      <input type="hidden" value="/V5/Code/DocumentUpload.asp" name="vNext" /> -->
          <input type="hidden" value="/V5/Code/Error.asp?vErr=Upload Successful!&vReturn=DocumentUpload.asp" name="vNext" />

        </form>
      </td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
