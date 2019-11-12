<%
  Dim vSset_Group, vSset_Lang, vSset_Id, vSset_ModIds
  Dim vSset_Eof

  '____ Sset ________________________________________________________________________

  '...Get SkillSet RecordSet
  Sub sGetSset_Rs  (vL, vG)
    vSql = "SELECT * FROM Sset WHERE Sset_Lang = '" & vL & "' "

    If vG <> "X" And vG <> "*" Then 
      vSql = vSql & " AND Sset_Group = '" & vG & "'"
    End If

    If vG = "X" Then 
      vSql = vSql & " ORDER BY Sset_Id, Sset_Lang "
    Else
      vSql = vSql & " ORDER BY Sset_Group, Sset_Lang, SSet_Id "
    End If

'   sDebug 
    sOpenDbBase
    Set oRsBase = oDbBase.Execute(vSql)
  End Sub  


  Sub sReadSset
    vSset_Group   = oRsBase("Sset_Group")
    vSset_Lang    = oRsBase("Sset_Lang")
    vSset_Id      = oRsBase("Sset_Id")
    vSset_ModIds  = oRsBase("Sset_ModIds")
  End Sub


  Sub sInitSset
    Dim aSkillSet, vSql1, vSql2, vSql3
    '...read Mods
    sOpenDbBase
    sOpenDbBase2
    '...first clear out all  
    vSql = "TRUNCATE TABLE Sset"
    Set oRsBase = oDbBase.Execute(vSql)    
    '...now build up the new Ssetl sets
    vSql = "Select * FROM Mods WHERE LEN(Mods_SkillSet) > 0"
    Set oRsBase = oDbBase.Execute(vSql)    
    Do While Not oRsBase.Eof 
      sReadMods
      aSkillSet = Split(vMods_SkillSet, "::")
      For j = 0 To Ubound(aSkillSet)
        vSset_Id = fUnquote(aSkillSet(j))
        vSql1 = "SELECT Sset_ModIds FROM Sset WHERE Sset_Group = '" & Left(vMods_Id, 1) & "' AND Sset_Lang = '" & Ucase(Right(vMods_Id, 2)) & "' AND Sset_Id = '" & vSset_Id & "'"
        Set oRsBase2 = oDbBase2.Execute(vSql1)    
  
        '...insert or update record 
        If oRsBase2.Eof Then 
          vSql2 = "INSERT INTO Sset (Sset_Group, Sset_Lang, Sset_Id, Sset_ModIds) VALUES ('" & Left(vMods_Id, 1) & "', '" & Ucase(Right(vMods_Id, 2)) & "', '" & vSset_Id & "', '" & Ucase(vMods_Id) & "')"
          oDbBase2.Execute(vSql2)
        Else
          vSset_ModIds = oRsBase2("Sset_ModIds")
          If Instr(vSset_ModIds, Ucase(vMods_Id)) = 0 Then
            vSql3 = "UPDATE Sset SET Sset_ModIds = '" & vSset_ModIds & " " & Ucase(vMods_Id) & "' WHERE Sset_Group = '" & Left(vMods_Id, 1) & "' AND Sset_Lang = '" & Ucase(Right(vMods_Id, 2)) & "' AND Sset_Id = '" & vSset_Id & "'" 
            oDbBase2.Execute(vSql3)
          End If
        End If  
      Next 
      oRsBase.MoveNext
    Loop
  
    Set oRsBase  = Nothing
    Set oRsBase2 = Nothing
    sCloseDbBase    
    sCloseDbBase2
  
  End Sub  
  
  
  
  
  
  
  
  

%>