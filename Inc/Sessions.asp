<% 
  Dim vTest
  vTest = True
  vTest = False
  If vTest Then
    Dim vSession1, vSession2
    sDisplaySessions  
    vSession1 = fStoreSessions
    Response.Write vSession1
    sKillSessions
    sReCreateSessions (vSession1)
    sDisplaySessions 
  End If


  Sub sDisplaySessions
    Dim vVar
    For Each vVar In Session.Contents
      Response.Write "Session(" & vVar & ")=" & Session(vVar) & "<br>"
    Next
  End Sub


  Function fStoreSessions
    Dim vVar
    fStoreSessions = ""
    For Each vVar In Session.Contents
      '...capture all "text" values - ie ignore arrays and objects, etc 
      If vVar <> "HostDbPwd" And VarType(Session(vVar)) <= 8 Then 
        fStoreSessions = fStoreSessions & "&" & vVar & "=" & Server.UrlEncode(fOkValue(Session(vVar)))
      End If     
    Next
    fStoreSessions = Mid(fStoreSessions, 2)
  End Function


  Function fSessionsJS
    Dim vVar
    fStoreSessionsJS = ""
    For Each vVar In Session.Contents
     If vVar <> "HostDbPwd" And VarType(Session(vVar)) < 16 Then 
       fStoreSessions = fStoreSessions & "&" & vVar & "=" & Server.UrlEncode(Session(vVar))
     End If     
    Next
    fStoreSessions = Mid(fStoreSessions, 2)
  End Function


  Sub sReCreateSessions (vSession1)
    Dim aVar1, aVar2
    aVar1 = Split(vSession1, "&")
    For i = 0 To Ubound(aVar1)
      aVar2 = Split(aVar1(i), "=")
      Session(aVar2(0)) = fUrlDecode(aVar2(1))
    Next
  End Sub


  Sub sKillSessions
    Session.Contents.RemoveAll()
  End Sub


	Function fUrlDecode(sConvert)
    Dim aSplit, sOutput, I
    fUrlDecode = ""
    If IsNull(sConvert) Then Exit Function
    If sConvert = "" Then Exit Function
    sOutput = Replace(sConvert, "+", " ")
    aSplit = Split(sOutput, "%")           ' next convert %hexdigits to the character
    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
        Chr("&H" & Left(aSplit(i + 1), 2)) &_
        Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If	
    fUrlDecode = sOutput
	End Function
%> 


 