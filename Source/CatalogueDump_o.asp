<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<% 
  Dim vPromo, vCatlT, vCatlN, vProgS, vProgN, vProgD, vProgL, vProgM, vProgE, vModsN, vModsI, vModsD, vFormt, vOrder

  vCatlT = Request("vCatlT")
  vCatlN = Request("vCatlN")
  vProgS = Request("vProgS")
  vProgN = Request("vProgN")
  vProgD = Request("vProgD")
  vProgL = Request("vProgL")
  vProgM = Request("vProgM")
  vModsN = Request("vModsN")
  vModsI = Request("vModsI")
  vModsD = Request("vModsD")
  vPromo = Request("vPromo")
  vFormt = Request("vFormt")
  vOrder = Request("vOrder")
%>

<html>

<head>
  <title>CatalogueDump_o</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>

  <p style="text-align:center; padding:20px;">
    <input onclick="javascript: history.back(1)" type="button" value="Return" class="button">
  </p>

  <table class="table">

    <%      
      Dim vCatlTitle, vProgTitle, vCatlPrev, vProgPrev, vSpace

      sOpenDb
      vSql =        " SELECT ca.Catl_CustId, ca.Catl_Title, ca.Catl_Programs, pr.Prog_Id, pr.Prog_Title1, pr.Prog_Desc, pr.Prog_Length, pr.Prog_Mods, " 
      vSql = vSql & " pr.Prog_Exam, pr.Prog_Assessment, pr.Prog_AssessmentCert" 

      If vModsN = "y" Or vModsD = "y" Or vModsI = "y" Then
      vSql = vSql & " , mo.Mods_Id, mo.Mods_Title, mo.Mods_Outline "
      End If
      
      vSql = vSql & " FROM V5_Vubz.dbo.Catl AS ca"
      vSql = vSql & " INNER JOIN V5_Base.dbo.Prog AS pr ON CHARINDEX(pr.Prog_Id, ca.Catl_Programs) > 0 "

      If vModsN = "y" Or vModsD = "y" Or vModsI = "y" Then
      vSql = vSql & " INNER JOIN V5_Base.dbo.Mods AS mo ON CHARINDEX(' ' + mo.Mods_Id, ' ' + pr.Prog_Mods) > 0 OR (pr.Prog_Assessment = mo.Mods_Id) "
      End If

      vSql = vSql & " WHERE (ca.Catl_CustId = '" & svCustId & "') AND (ca.Catl_Active = 1)"
'     vSql = vSql & "  AND (pr.Prog_Id = 'P2072EN') "


      If vCatlT <> "0" Then
      vSql = vSql & " AND  (' ' + CHARINDEX(CAST(ca.Catl_No AS varchar), ' " & vCatlT & " ') > 0)"
      End If
      
      vSql = vSql & " ORDER BY ca.Catl_Title,  pr.Prog_Title1"


      vSpace = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
      vCatlPrev = ""
      vProgPrev = ""
      
'     sDebug

      Set oRs = oDb.Execute(vSQL)    
      Do While Not oRs.Eof 
        vCatlTitle = oRs("Catl_Title")
        If vPromo = "y" Then
          i = Instr(vCatlTitle, "<") 
          If i > 0 Then 
            vCatlTitle = Left(vCatlTitle, i - 1)
          End If
        
        End If

        '...ignore inactive programs "P1194EN~9999~9999~1~90"
        i = Instr(oRs("Catl_Programs"), oRs("Prog_Id"))
        If Mid(oRs("Catl_Programs"), i + 8, 4) <> "9999" Then

          vProgTitle = oRs("Prog_Title1")
          If vPromo = "y" Then
            i = Instr(vProgTitle, "<") 
            If i > 0 Then 
              vProgTitle = Left(vProgTitle, i - 1) 
            End If
          End If
  
          If vCatlTitle <> vCatlPrev Then
            vCatlPrev = vCatlTitle
            If vCatlN = "y" Then
    %>
    <tr>
      <td class="c1">
        <%=vCatlTitle%>
        <% If vProgs="y" Then %>
        <blockquote>
          <span style="font-size: small"><%=oRs("Catl_Programs")%></span>
        </blockquote>
        <% End If %>
      </td>
    </tr>
    <%  
            End If
          End If 
  
          If vProgTitle <> vProgPrev Then
            vProgPrev = vProgTitle
    %>
    <tr>
      <td class="c2">
        <blockquote>
          <%=vProgTitle & " (" & oRs("Prog_Id")%>)
        </blockquote>
      </td>
    </tr>

    <%  
            If vProgD = "y" Then
    %>
    <tr>
      <td class="c3">
        <blockquote>
          <blockquote>
            <%=oRs("Prog_Desc")%>

            <% If vProgE = "y" Then %>
              <% If Len(oRs("Prog_Assessment")) > 0 Or Lcase(oRs("Prog_Exam")) <> "n" Then %>
            <br>
            <br>
            An Examination is available with this Program (<%=oRs("Prog_Assessment")%>).
              <% End If %>

              <% If Len(oRs("Prog_AssessmentCert"))> 0 Then %>
            <br>
            <br>
            An Certificate of Completion is available with this Program.
              <% End If %>
            <% End If %>

            <% If vProgD = "n" Then %>
            <br>
            <br>
            Estimated Program Length: <%=oRs("Prog_Length")%> hrs.
            <% End If %>

            <% If vProgM = "y" Then %>
            <br>
            <br>
            Program contains <%=Ubound(Split(oRs("Prog_Mods"))) + 1  + fIf(Len(oRs("Prog_Assessment")) >0 , 1, 0) %> module(s).
            <% End If %>
          </blockquote>
        </blockquote>
      </td>
    </tr>
    <%  
            End If 
    %>




   <%
          '...  get module info (in sequence) for this program
          sOpenDb2
          vSql =        " SELECT mo.Mods_Id, mo.Mods_Title, mo.Mods_Outline "
          vSql = vSql & " FROM "
	        vSql = vSql & "   V5_Base.dbo.Mods AS mo INNER JOIN "
	        vSql = vSql & "   apps.dbo.Split ((SELECT Prog_Mods + ' ' + Prog_Assessment FROM V5_Base.dbo.Prog WHERE Prog_Id = '" & oRs("Prog_Id") & "'), ' ') ON mo.Mods_Id = strval "
'         sDebug
          Set oRs2      = oDb2.Execute(vSQL)       
          Do While Not oRs2.Eof            
            vMods_Id      = oRs2("Mods_Id")
            vMods_Title   = oRs2("Mods_Title")
            vMods_Outline = oRs2("Mods_Outline")
      
    %>

    <%  
            If vModsN = "y" Then
    %>
    <tr>
      <td class="c4">
        <blockquote>
          <blockquote>
            <blockquote>
              <%=vMods_Title & " (" & vMods_Id & ")" %>
            </blockquote>
          </blockquote>
        </blockquote>
      </td>
    </tr>
    <%  
            End If 
    %>


    <%  
            If vModsI = "y" Then
    %>
    <tr>
      <td class="c4">
        <blockquote>
          <blockquote>
            <blockquote>
              <%=vMods_Id%>
            </blockquote>
          </blockquote>
        </blockquote>
      </td>
    </tr>
    <%  
            End If 
    %>

    
    <%  
            If vModsD = "y" Then
    %>
    <tr>
      <td>
        <blockquote>
          <blockquote>
            <blockquote>
              <blockquote>
                <%=vMods_Outline%>
              </blockquote>
            </blockquote>
          </blockquote>
        </blockquote>
      </td>
    </tr>
    <%  
            End If 
    %>


    <%  
            oRs2.MoveNext
          Loop
          Set oRs2 = Nothing
          sCloseDb2 
    
        End If 
    %>


    <%
        End If
        oRs.MoveNext
      Loop
    %>

  </table>

  <%
      Set oRs = Nothing
      sCloseDb    
  %>

  <p style="text-align:center; padding:20px;">
    <input onclick="javascript: history.back(1)" type="button" value="Return" class="button">
  </p>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
