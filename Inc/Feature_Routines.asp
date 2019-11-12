<%
  Dim vFeatureFolderA, vFeatureFolderR, aFeatures

  '...define the absolute/virtual folder for the text file object and the relative folder from the Root folder for the browser
  vFeatureFolderR = "..\Features\"
  vFeatureFolderA = Server.MapPath("/V5/Features")

  Sub sGetFeatures(vFeatureClass, vTotalNo, vStartNo, vListNo)
   
    If Not svSecure Then Exit Sub
    '...the full feature class requires the cluster no (C0002EN) less the "C" part
    '...for the banner or left side, get an array of values
    '   but not for centre of right columns because we want to use the same array as the "L" side
    If vFeatureClass = "B" Or vFeatureClass = "L" Then
      '...If Left Column look for a "C" column feature
      If vFeatureClass = "L" Then
        vFeatureClass = "C"
      End If         
      vFeatureClass = vFeatureClass & Mid(svCustCluster, 2)
      '...setup and display
      aFeatures = fSetupFeatureArray (vFeatureFolderA, vFeatureClass, vTotalNo)
      vFeatureClass = vFeatureClass & Mid(svCustCluster, 2)
    Else 
      '...must be either a "C"entre or "R"ight column so make "C"olumn Feature
      vFeatureClass = "C" & Mid(svCustCluster, 2)
    End If
    '...display the features
    sDisplayFeatureArray vFeatureFolderR, vStartNo, vListNo
  End Sub


  Function fSetupFeatureArray (vPath, vClass, vNo)
    Dim oFileSys, oFolder, oFile
    Dim aAllFiles, aReturnFiles()
    Dim vFiller, vRandom, vRepeat, i, j
    Randomize Timer
    ReDim aReturnFiles(vNo - 1)
    Set oFileSys = CreateObject("Scripting.FileSystemObject")
    Set oFolder = oFileSys.GetFolder(vPath)
    '...Gather all files names into array duplicating names according to rank
    For Each oFile In oFolder.Files
      If Left(oFile.Name, Len(vClass)) = vClass And Len(oFile.Name) > 18 Then
        For vFiller = 1 To Mid(oFile.Name, 14, 1)
          If VarType(aAllFiles) = vbEmpty Then
            ReDim aAllFiles(0)
          Else
            ReDim Preserve aAllFiles(UBound(aAllFiles) + 1)
          End If
          aAllFiles(UBound(aAllFiles)) = oFile.Name
        Next
      End If
    Next
    '...Parse list and pull out vNo random names
    For i = 0 To (vNo - 1)
      vRepeat = False
      Do
        vRandom = Int((UBound(aAllFiles) + 1) * Rnd)
        '...Make sure we have a unique name
        vRepeat = False
        For j = 0 To i - 1
          If aReturnFiles(j) = aAllFiles(vRandom) Then
            vRepeat = True
            Exit For
          End If
        Next
        If (Not vRepeat) Or (i = 0) Then
          aReturnFiles(i) = aAllFiles(vRandom)
          Exit Do
        End If
      Loop
    Next
    fSetupFeatureArray = aReturnFiles
  End Function


  Sub sDisplayFeatureArray (vFeatureFolderR, vStartNo, vNo)
    For i = vStartNo - 1 To vStartNo + vNo - 1 - 1
      Server.Execute vFeatureFolderR & aFeatures(i)
    Next
  End Sub

%>