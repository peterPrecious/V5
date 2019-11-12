<%
  Dim vKeys_No, vKeys_AcctId, vKeys_TskHNo, vKeys_MembNo, vKeys_Posted
  Dim vKeys_Type, vKeys_Script

  bDebug = False

  '...get the script type ("none", "task", "site") and file address of the script
  If Len(Session("Keys_Type")) = 0 Then
    sGetScript
  Else
    vKeys_Type   = Session("Keys_Type")
    vKeys_Script = Session("Keys_Script")
  End If

  If bDebug Then Response.Write "<br>Keys_Type: " & vKeys_Type & ", vKeys_Script: " & vKeys_Script



  '____ Keys  ________________________________________________________________________

  '...If there is a key for this task or site then it is not locked, else it is deemed locked
  '   Note: if there is no script there can be no keys since the script creates the keys

  Function fIsLocked (vTskHNo, vMembNo)
    bDebug = False
    If vKeys_Type = "none" Then 
      fIsLocked = True
      fIsLocked = False '...not sure why previous line exists...
      Exit Function
    ElseIf vKeys_Type = "task" Then '...check for task keys
      vSql = "SELECT * FROM Keys WHERE Keys_AcctId = '" & svCustAcctId & "' AND Keys_TskHNo = " & vTskHNo & " AND Keys_MembNo = " & vMembNo
    Else     '...check for sitewise keys 
      vSql = "SELECT * FROM Keys WHERE Keys_AcctId = '" & Left(svCustId, 4) & "' AND Keys_MembNo = " & vMembNo
    End If
    sOpenDb2
    Set oRs2 = oDb2.Execute(vSql)
    If oRs2.Eof Then 
      fIsLocked = True
    Else  
      fIsLocked = False
    End If  
    Set oRs2 = Nothing
    sCloseDb2
    If bDebug Then Response.Write "<br>" & vSql & " - IsLocked: " & fIsLocked
  End Function


  '...this deletes a key(s) that may exist to open a single task vTskHNo (if vTskHNo is numeric) or all locked tasks in the site (if vTskNo is alpha)
  Sub sLock (vTskHNo, vMembNo)
    If vKeys_Type = "none" Then 
      Exit Sub  
    ElseIf vKeys_Type = "task" Then '...check for task keys
      vSql = "DELETE FROM Keys WHERE Keys_AcctId = '" & svCustAcctId & "' AND Keys_TskHNo = " & vTskHNo & " AND Keys_MembNo = " & vMembNo
    Else
      vSql = "DELETE FROM Keys WHERE Keys_AcctId = '" & Left(svCustId, 4) & "' AND Keys_MembNo = " & vMembNo
    End If
    sOpenDb2
    oDb2.Execute(vSql)
    sCloseDb2
    If bDebug Then Response.Write "<br>" & vSql
  End Sub


  '...this creates a key that will unlock all or a specific task in this account
  Sub sUnLock (vTskHNo, vMembNo)
    vSql = "INSERT INTO Keys"
    vSql = vSql & "(Keys_AcctId, Keys_TskHNo, Keys_MembNo)"

    If vKeys_Type = "none" Then 
      Exit Sub  
    ElseIf vKeys_Type = "task" Then '...check for task keys
      vSql = vSql & " VALUES ('" & svCustAcctId & "', " & vTskHNo & ", " & vMembNo & ")"
    Else
      vSql = vSql & " VALUES ('" & Left(svCustId, 4) & "', 0, " & vMembNo & ")"
    End If
    sOpenDb2
    oDb2.Execute(vSql)
    sCloseDb2
    If bDebug Then Response.Write "<br>" & vSql
  End Sub



  '...check if there are any unlocking scripts in the repository
  '   store the result in a session variable
  '   NOTE look for the repository first in the Virtual Webs (live) then in Active Webs folder (dev)
  Sub sGetScript
    Dim oFs, vRoot, vWebs, vFolder, vFile
    Set oFs = CreateObject("Scripting.FileSystemObject")

    vKeys_Type = "none"  '...assume no script unless we find one (either a task level script or a site wide script)

    vRoot   = Server.MapPath("\V5") 
    If bDebug Then Response.Write "<br>Root:&nbsp;&nbsp;" & vRoot   & " - " & oFs.FolderExists(vRoot)
    vWebs   = Left(vRoot, Len(vRoot) - 10)
    If bDebug Then Response.Write "<br>Webs:&nbsp;&nbsp;" & vWebs   & " - " & oFs.FolderExists(vWebs)
    vFolder = vWebs & "\Virtual\Repository\V5_Vubz\0000\Keys"
    If bDebug Then Response.Write "<br>Folder:&nbsp;&nbsp;" & vFolder & " - " & oFs.FolderExists(vFolder)

    If Not oFs.FolderExists(vFolder) Then
      vFolder = vWebs & "\Active\V5\Repository\V5_Vubz\0000\Keys"
      If bDebug Then Response.Write "<br>Folder:&nbsp;&nbsp;" & vFolder & " - " & oFs.FolderExists(vFolder)
    End If

    '...look to run an unlock script for a key for this specific site
    vFile  = vFolder & "\" & svCustAcctId & ".asp"
    If bDebug Then Response.Write "<br>File:&nbsp;&nbsp;" & vFile & " - " & oFs.FileExists(vFile)
    If oFs.FileExists(vFile) Then
      vKeys_Type = "task"
      vFile  = "/V5/Repository/V5_Vubz/0000/Keys/" & svCustAcctId & ".asp"
      '...use relative URLs rather then just vFile (go figure)
      If bDebug Then Response.Write "<br>Script available at: " & vFile
'     Server.Execute vFile

    '...look to run an unlock script for a key for this group of sites
    Else 
      vFile  = vFolder & "\" & Left(svCustId, 4) & ".asp"
      If bDebug Then Response.Write "<br>File:&nbsp;&nbsp;" & vFile & " - " & oFs.FileExists(vFile)
      If oFs.FileExists(vFile) Then
        vKeys_Type = "site"
        vFile  = "/V5/Repository/V5_Vubz/0000/Keys/" & Left(svCustId, 4) & ".asp"
        '...use relative URLs rather then just vFile (go figure)
        If bDebug Then Response.Write "<br>Script available at: " & vFile
'       Server.Execute vFile
      End If

   End If

   Set oFs = Nothing
   vKeys_Script = vFile

   Session("Keys_Type")   = vKeys_Type
   Session("Keys_Script") = vKeys_Script
    
  End Sub 



  Sub sUnlocks
    If bDebug Then Response.Write "<br>About to Run Unlocking Script"
    If vKeys_Type <> "none" Then
      If bDebug Then Response.Write "<br>Script starting: " & vKeys_Script
      Server.Execute vKeys_Script
      If bDebug Then Response.Write "<br>Script ending: " & vKeys_Script
    Else
      If bDebug Then Response.Write "<br>No Script available"
    End If
  End Sub














%>