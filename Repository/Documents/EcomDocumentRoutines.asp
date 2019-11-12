<%
  '...Leave in Source/Code for easy of path finding (from GetPath.asp)

  '...vDocuments is the folder name containing this set of documents (Manuals), ie "MultiUserManual"
  '...vDocument is the file name of the document (Manual), ie "VUBZ_MultiUserManual_FR.pdf" - this is always held in a folder "MultiUserManual"

  Dim vManualAudit
  vManualAudit = True
  vManualAudit = False

  '...use next line for testing only
' Response.Write svMultiUserManual
  

  Function fListDocuments (vDocuments)
    fListDocuments = "<br>"
    Dim oFs, oFolder, oFiles, vFile, vFolder, vDelUrl
    vFolder = svMultiUserManual
    Set oFs = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFs.GetFolder(vFolder)
    Set oFiles = oFolder.Files
    For Each vFile in oFiles    
      If Right(vFile.Name, 4) = ".pdf" Then
        vDelUrl = "location.href=" & Chr(34) & "UploadMultiUserManual.asp?vDelete=" & vFile.Name & Chr(34)    
        fListDocuments = fListDocuments & "<input onclick='" & vDelUrl & "' class='button'  style='margin:2px; height: 20px' type='button' value='Delete' name='bDelete'>&nbsp;&nbsp;"
        fListDocuments = fListDocuments & "<a target='_blank' href='/V5/Repository/Documents/" & vDocuments & "/" & vFile.Name & "'>" & vFile.Name & "</a>"
        fListDocuments = fListDocuments & "<br>"
      End If
    Next
    Set oFs     = Nothing
    Set oFolder = Nothing
    Set oFiles  = Nothing
  End Function


  '...this will get the appropriate document in the vDocuments folder
  '   will look for 8 char CustId (ie ABCD1234), then 4 Char CustId (ABCD), then VUBZ, then Language

  Function fGetDocument (vDocuments)
    fGetDocument = ""
    Dim oFs, oFolder, oFiles, vFile, vFolder, aFileName
    Set oFs = CreateObject("Scripting.FileSystemObject")
    If Len(svMultiUserManual) = 0 Then Exit Function
    vFolder = svMultiUserManual
    Set oFolder = oFs.GetFolder(vFolder)
    Set oFiles = oFolder.Files

    '...look for 8 char value
    If fGetDocument = "" Then
      For Each vFile in oFiles
        If Right(vFile.Name, 4) = ".pdf" Then
          aFileName = Split(Ucase(vFile.Name), "_")
          If svCustId = aFileName(0) And svLang = Left(aFileName(2), 2) Then
            fGetDocument = vFile.Name
            Exit For
          End If  
        End If
      Next
    End If

    '...look for 4 char value
    If fGetDocument = "" Then
      For Each vFile in oFiles
        If Right(vFile.Name, 4) = ".pdf" Then
          aFileName = Split(Ucase(vFile.Name), "_")
          If Left(svCustId, 4) = aFileName(0) And svLang = Left(aFileName(2), 2) Then
            fGetDocument = vFile.Name
            Exit For
          End If  
        End If
      Next      
    End If
    
    '...look for 4 char value
    If fGetDocument = "" Then
      For Each vFile in oFiles
        If Right(vFile.Name, 4) = ".pdf" Then
          aFileName = Split(Ucase(vFile.Name), "_")
          If "VUBZ" = aFileName(0) And svLang = Left(aFileName(2), 2) Then
            fGetDocument = vFile.Name
            Exit For
          End If  
        End If
      Next
    End If    

    If fGetDocument <> "" Then
      fGetDocument = "/V5/Repository/Documents/" & vDocuments & "/" & fGetDocument 
    End If

    Set oFs     = Nothing
    Set oFolder = Nothing
    Set oFiles  = Nothing
  End Function


  Sub sDeleteDocument (vDocument)
    Dim oFs, vFolder, vDocuments
    vFolder = svMultiUserManual
    vDocuments = "MultiUserManual"
    vDocument = vFolder & "\" & vDocument
    If vManualAudit Then Response.Write "<p>Deleting " & vDocument & " (...inactive...)"
    Set oFs = CreateObject("Scripting.FileSystemObject")
    oFs.DeleteFile(vDocument)
  End Sub

%>

