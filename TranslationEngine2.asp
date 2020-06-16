<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->

<% 
  Dim vSource, vDestination, vGoodFiles, vBadFiles, vGoodCnt, vBadCnt
  Dim oFs, oFolder, oFiles, oInp, oOut
  Const ForReading = 1, ForWriting = 2

  vSource       = "\V5\Source"
  vDestination  = "\V5\Code"

  vGoodFiles = "" : vGoodCnt = 0
  vBadFiles  = "" : vBadCnt  = 0

  sTranslate 

  '--------------  Translation Funcitons -------------------------

  Sub sTranslate ()

    Dim vNoFiles, vFileNo, vFile, vTemp, vLine, vLineSave, vOk, aFile, vPhraEn, vPhraNo, vPhraHidden, vSelectPages, vDbName, vDbPwd, vHeaderStart, vHeaderEnd

    vSelectPages  = Request("vSelectPages")

    vNoFiles = 0
    vGoodCnt = 0 

    '...get all files in the Source folder and put into an array
    Set oFs = CreateObject("Scripting.FileSystemObject")   
    Set oFolder = oFs.GetFolder(Server.MapPath(vSource))
    Set oFiles = oFolder.Files

    '...get all files in web
    For Each vFile in oFiles
      ReDim Preserve aFileIn (vNoFiles)
      aFileIn(vNoFiles) = vFile.Name
      vNoFiles = vNoFiles + 1
    Next 

    '...go through the array and extract what we need to translate
    If vNoFiles > 0 Then 
      '...start translating
      For vFileNo = 0 to Ubound(aFileIn)

      	vPage = aFileIn(vFileNo)    

        '...valid page to translate?
        If vSelectPages = "all" Then 
          vOk = True
        ElseIf Instr(vSelectPages, vPage) > 0 Then 
          vOk = True
        Else  
          vOk = False
        End If

      	'...open input file 
        If vOk Then

      	  Set oInp = oFs.OpenTextFile(Server.MapPath(vSource) & "\" & vPage, ForReading, True)
      	  vLine = oInp.ReadAll
          Set oInp = Nothing    

          '...ensure it's valid
          If fFileOk (vLine, vPage) Then

            '...if ok then open output file
        	  Set oOut = oFs.OpenTextFile(Server.MapPath(vDestination) & "\" & vPage, ForWriting, True)
            vGoodFiles = vGoodFiles & vPage & " "
            vGoodCnt = vGoodCnt + 1

            '...only check .asp and .js files
            If Right(vPage, 4) = ".asp" Or Right(vPage, 3) = ".js" Then
            
              '...convert any links to their language equivalents
          	  vLine = Replace(vLine,"_EN.asp","_<##=svLang##>.asp")
          	  vLine = Replace(vLine,"_EN.gif","_<##=svLang##>.gif")
          	  vLine = Replace(vLine,"_EN.jpg","_<##=svLang##>.jpg")
          	  vLine = Replace(vLine,"_EN.swf","_<##=svLang##>.swf")
    
    
              '...check for all html phrases, ie: <br><!--[[-->Eat my shorts.<!--]]-->
              vPhraHidden = False
        	    Do While Instr(vLine, "[[") > 0
                '...isolate phrase
        	      i = Instr(vLine, "[[")     '...start of phrase
        	      j = Instr(vLine, "]]")     '...end of phrase
        	      vPhraEn = Mid(vLine, i + 5, j - i - 9)
        	      vPhraNo = Right("000000" & fPhraNo(vPhraEn), 6)    	      
                '...reformat line using function to return and update phrase table             
         	      vLine = Left(vLine, i - 5) & "<!--webbot bot='PurpleText' PREVIEW='" & vPhraEn & "'-->" & "<##=fPhra(" & vPhraNo & ")##>" & Mid(vLine, j + 5)
        	    Loop
    
    
              '...check for vb script phrases, ie: vMsg="<!--{{-->Eat my shorts.<!--}}-->"
              vPhraHidden = True
        	    Do While Instr(vLine, "{{") > 0
                '...isolate phrase
        	      i = Instr(vLine, "{{")     '...start of phrase
        	      j = Instr(vLine, "}}")     '...end of phrase
        	      vPhraEn = Mid(vLine, i + 5, j - i - 9)
                '...reformat line using function to return and update phrase table
         	      vLine = Left(vLine, i - 6) & "fPhraH(" & Right("000000" & fPhraNo(vPhraEn), 6) & ")" & Mid(vLine, j + 6)
        	    Loop
    
    
              '...check for java script phrases, ie: var vPhrase1 = "/*--{[--*/Eat my shorts./*--]}--*/"
              vPhraHidden = True
        	    Do While Instr(vLine, "{[") > 0
                '...isolate phrase
        	      i = Instr(vLine, "{[")     '...start of phrase
        	      j = Instr(vLine, "]}")     '...end of phrase
        	      vPhraEn = Mid(vLine, i + 6, j - i - 10)
                '...reformat line using function to return and update phrase table
         	      vLine = Left(vLine, i - 6) & """<##=fPhraH(" & Right("000000" & fPhraNo(vPhraEn), 6) & ")##>""" & Mid(vLine, j + 7)
        	    Loop
  
        	    '...twig scripting values
        	    vLine = Replace(vLine, "##", "%")
      	    
      	    End If
      	    
         	  '...write out lines
         	  oOut.WriteLine vLine & vbCrLf
          	oOut.Close  
            Set oOut = Nothing

          End If
  
        End If    

      Next  

    End If  

    
    '...if we have some good files, then assign pages to phrases, first clean out all previous page references
    If vGoodCnt > 0 Then
  
      sPhraClean
  
      '...now add the page names that any phrase refers to (unless they are "badfiles")
      For vFileNo = 0 to Ubound(aFileIn)
  
      	vPage = aFileIn(vFileNo)

        If Instr(vBadFiles, vPage) = 0 Then
  
          vFile = Server.MapPath(vDestination) & "\" & vPage
          Set oInp = oFs.OpenTextFile(vFile, ForReading, True)

          If Not oInp.AtEndOfStream Then '...ensure file exists
            vLine = oInp.ReadAll
            vLineSave = vLine
      
            '...check for any "visible" function calls, ie: <br>fPhra(999999)
            Do While Instr(vLine, "fPhra(") > 0
              i = Instr(vLine, "fPhra(") 
              vPhraNo = Mid(vLine, i + 6, 6)
              sPhraPages vPhraNo, vPage
              vLine = Mid(vLine, i + 10)
            Loop
      
            '...check for any "hidden" function calls, ie: <br>fPhraH(999999)
            vLine = vLineSave
            Do While Instr(vLine, "fPhraH(") > 0
              i = Instr(vLine, "fPhraH(") 
              vPhraNo = Mid(vLine, i + 7, 6)
              sPhraHPages vPhraNo, vPage
              vLine = Mid(vLine, i + 10)
            Loop
            
          End If
          
        End If
        
      Next

    End If

    Set oFs  = Nothing
    Set oInp = Nothing
  End Sub


  Function fFileOk(vLine, vPage)

    fFileOk = False
    '...see if there are proper pairs of either [[]] or {{}} or {[]}
    If Ubound(Split(vLine, "--[[--")) <> Ubound(Split(vLine, "--]]--")) Then 
      vBadFiles = vBadFiles & vPage & " ( [[...]] )~"
      vBadCnt = vBadCnt + 1
    ElseIf Ubound(Split(vLine, "--{{--")) <> Ubound(Split(vLine, "--}}--")) Then
      vBadFiles = vBadFiles & vPage & " ( {{...}} )~"
      vBadCnt = vBadCnt + 1
    ElseIf Ubound(Split(vLine, "--{[--")) <> Ubound(Split(vLine, "--]}--")) Then
      vBadFiles = vBadFiles & vPage & " ( {[...]} )~"
      vBadCnt = vBadCnt + 1
    Else
      fFileOk = True
    End If
  End Function


  '...get Phrase no, if not on file add
  Function fPhraNo (vPhraEn)
    Dim vPhra_No, vPhra_EN, vPhra_FR, vPhra_ES, vPhra_Len, vPhra_Pages
    sOpenDb
    vSql = "SELECT Phra_No FROM Phra WHERE Phra_EN = '" & fUnquote(vPhraEn) & "'"    
    Set oRs = oDb.Execute(vSql)
    If oRs.Eof Then 
      vSql = " SET NOCOUNT ON " _
           & " INSERT INTO Phra (Phra_EN) VALUES ('" & fUnquote(vPhraEn) & "')" _
           & " SELECT Phra_No=@@IDENTITY" _
           & " SET NOCOUNT OFF"
      Set oRs = oDb.Execute(vSql)
'     sDebug "vSql", vSql
    End If   
    fPhraNo = oRs("Phra_No")
    Set oRs = Nothing
    sCloseDb
  End Function


  '...get Phrase no, if not on file add
  Function fPhra (vPhraNo)
    Dim vPhra_EN
    sOpenDb
    vSql = "SELECT Phra_EN FROM Phra WHERE Phra_No = " & vPhraNo
    Set oRs = oDb.Execute(vSql)
    If oRs.Eof Then 
      fPhra = ""
    Else
      fPhra = oRs("Phra_EN")
    End If      
    Set oRs = Nothing
    sCloseDb
  End Function


  '...update "visible" Phrase pages
  Sub sPhraPages (vPhraNo, vPhraPage)
    sOpenDb
    vSql = "UPDATE Phra SET Phra_Pages = '" & vPhraPage & "' + ' ' + Phra_Pages WHERE Phra_No = " & vPhraNo & " AND CHARINDEX('" & vPhraPage & "', Phra_Pages) = 0"
    oDb.Execute(vSql)     
    sCloseDb
  End Sub
  

  '...update "hidden" Phrase pages
  Sub sPhraHPages (vPhraNo, vPhraPage)
    sOpenDb
    vSql = "UPDATE Phra SET Phra_Hidden = '" & vPhraPage & "' + ' ' + Phra_Hidden WHERE Phra_No = " & vPhraNo & " AND CHARINDEX('" & vPhraPage & "', Phra_Hidden) = 0"
    oDb.Execute(vSql)     
    sCloseDb
  End Sub  
  

  '...remove all page references as values will be re-entered
  Sub sPhraClean()
    sOpenDb
    vSql = "UPDATE Phra SET Phra_Pages = '', Phra_Hidden = ''" 
    oDb.Execute(vSql)     
    sCloseDb
  End Sub

%>

<html>

<head>
  <title>:: Translation Engine 2/2</title>
  <meta charset="UTF-8">
  <script src="Inc/jQuery.js"></script>
  <link href="Inc/Vubi2.css" rel="stylesheet" />
  <script src="Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <!--#include virtual = "V5/Inc/Shell_HiSolo.asp"-->

  <h1>Vubiz Translation Engine</h1>

  <table style="width: 600px; margin: auto;">
    <tr>
      <td style="text-align: center">

        <% If vGoodCnt = 1 Then %>
        <h2>The following file was translated:</h2>
        <% ElseIf vGoodCnt > 1 Then %>
        <h2>The following <%=vGoodCnt%> files were translated:</h2>
        <% End If %>

        <% If vGoodCnt > 0 Then %>

        <table style="width: 200px; margin: auto">
          <tr>
            <td><%=Replace(vGoodFiles, " ", "<br>")%></td>
          </tr>
        </table>

        <% End If %>

        <% If vBadCnt = 1 Then %>
        <h6>Note: The following file was not be translated<br>because of missing translation tags:</h6>
        <% ElseIf vBadCnt > 1 Then %>
        <h6>Note: The following files were not be translated<br>because of missing translation tags:</h6>
        <% End If %>

        <% If vBadCnt > 0 Then %>
        <div style="text-align: center">
          <table>
            <tr>
              <td style="white-space: nowrap"><%=Replace(vBadFiles, "~", "<br>")%></td>
            </tr>
          </table>
        </div>
        <% End If %>

        <br /><br />
        <input class="button" onclick="history.back(1)" type="button" value="Return" name="bReturn"><%=f10()%>
        <input class="button" onclick="location.href = '<%=svDomain%>/Default.asp?vCust=VUBZ2274&vId=<%=vPassword5%>&vSource=/V5/TranslationEngine1.asp'" type="button" value="Big Admin" name="bReturn">
      </td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
