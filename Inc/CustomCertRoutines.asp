<%
  Function fCustomCertOptions (vCustomCert)
    fCustomCertOptions = vbCrLf & "<option value=" & Chr(34) & Chr(34) & ">Not Used</option>"
    Dim oFs, oFolder, oSubFolders, vFolder, vFolderName, vSelected
    Set oFs = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFs.GetFolder(Server.MapPath("..\..\..\Virtual\Content\Assessments\CustomCerts"))
    Set oSubFolders = oFolder.SubFolders
    For Each vFolder in oSubFolders
      vFolderName = vFolder.Name 
      If Lcase(vFolderName) <> "components" And Lcase(vFolderName) <> "images" And Left(vFolderName, 1) <> "_" Then
        vSelected = fIf(vCustomCert = vFolderName, " selected", "")
        fCustomCertOptions = fCustomCertOptions & "<option value=" & Chr(34) & vFolderName & Chr(34) & vSelected & ">" & vFolderName & "</option>" & vbCrLf
      End If  
    Next
    Set oFs     = Nothing
    Set oFolder = Nothing
    Set oSubFolders  = Nothing
  End Function
%>
