<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_TskH.asp"-->
<!--#include virtual = "V5/Inc/Db_TskD.asp"-->
<!--#include virtual = "V5/Inc/Db_Crit.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<%
  Function fLevel(vLevel) 
    Select Case vLevel
      Case 1    : fLevel = "Learner"
      Case 2    : fLevel = "Learner +"
      Case 3    : fLevel = "Facilitator +"
      Case 4    : fLevel = "Manager +"
      Case 5    : fLevel = "Administrator"
      Case Else : fLevel = "Inactive"
    End Select
  End Function

  Function fCheckMark(i)
    If i = True Then
      fCheckMark = "<img border='0' src='../Images/Icons/Check.gif'>"
    Else  
      fCheckMark = "<font face='Verdana' size='1'>&nbsp;</font>" 
    End If
  End Function
  
  Function fDates (i)
    If Trim(i) = "-" Then
      fDates = " "
    Else
      fDates = i   
    End If     
  End Function

  Function fPassword(i)
    If Len(fOkValue(i)) = 0 Then
      fPassword = ""
    Else
      fPassword = String(Len(i), "*")
    End If
  End Function
%>


<html>

<head>
  <title>MyWorldFilters</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table class="table">
    <tr>
      <td colspan="10" style="text-align:center">
      <h1>My Learning Access Filters</h1>
      <h3>This shows the various restrictions that may apply to specific tasks.<br />Task Levels are represented by the level of indent.<To view any embedded Assets, click on the Minor Task Title.<br /><br /></h3>
      </td>
    </tr>
    <tr>
      <th class="rowshade" style="width:40%; text-align:left">Task Title</th>
      <th class="rowshade" style="width:05%;">Active?</th>
      <th class="rowshade" style="width:05%;">Dates</th>
      <th class="rowshade" style="width:05%; text-align:left">Level</th>
      <th class="rowshade" style="width:05%;">Lang</th>
      <th class="rowshade" style="width:05%; text-align:left">Group 1</th>
      <th class="rowshade" style="width:05%;">Group 2</th>
      <th class="rowshade" style="width:10%;">Cust <br>Ids</th>
      <th class="rowshade" style="width:10%; text-align:left">Learner<br>Ids</th>
      <th class="rowshade" style="width:10%; text-align:left">Password</th>
    </tr>

    <%
      Dim vTitle, vSpaces, vSpace, vGroup, vGroupJobs
      vTskH_AcctId = Request("vTskH_AcctId")
      vTskH_Id     = Request("vTskH_Id")
      vSpaces      = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"  _
                   & "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"  
  
      sGetTskH_Rs vTskH_AcctId, vTskH_Id
      
      Do While Not oRs.Eof
        sReadTskH    
  
        Select Case vTskH_Level
          Case 0
            vSpace = ""
            vTitle = "<b>" & fIf(Len(Trim(vTskH_Title)) = 0 Or (Left(Trim(vTskH_Title), 1) = "<" And Right(Trim(vTskH_Title), 1) = ">"), "[No Title]", fLeft(vTskH_Title, 48)) & "</b>"
          Case 1
            vSpace = Left(vSpaces, 24)
            vTitle = "<b>" & fIf(Len(Trim(vTskH_Title)) = 0 Or (Left(Trim(vTskH_Title), 1) = "<" And Right(Trim(vTskH_Title), 1) = ">"), "[No Title]", fLeft(vTskH_Title, 48)) & "</b>"
          Case 2 
            vSpace = Left(vSpaces, 48)
            vTitle = fIf(Len(Trim(vTskH_Title)) = 0 Or (Left(Trim(vTskH_Title), 1) = "<" And Right(Trim(vTskH_Title), 1) = ">"), "[No Title]", fLeft(vTskH_Title, 48))
      End Select

      vGroup = fCriteria(vTskH_Criteria)
      If vTskH_Criteria <> "0" Then 
        vGroupJobs = fCriteriaJobs(vTskH_Criteria)
        If Len(vGroupJobs) > 0 Then vGroup = vGroup & " (" & vGroupJobs & ")"
      End If
  %>
    <tr>
      <td>
        <% If vTskH_Level = 2 Then %>
        <%=vSpace%><a href="#" onclick="toggle('div<%=vTskH_No%>')"><%=vTitle%></a>
        <% Else %>
        <%=vSpace%><%=vTitle%>
        <% End If %>
        &nbsp; 
      </td>
      <td style="text-align:center"><%=fIf(vTskH_Active, "Y", "N")%></td>
      <td style="text-align:center" nowrap><%=fDates(fFormatDate(vTskH_DateStart) & " - " & fFormatDate(vTskH_DateStart))%>&nbsp; </td>
      <td><%=fLevel(vTskH_AccessLevel)%>&nbsp; </td>
      <td style="text-align:center"><%=vTskH_Lang%> </td>
      <td><%=vGroup%>&nbsp; </td>
      <td style="text-align:center"><%=vTskH_Group2%></td>
      <td style="text-align:center"><%=vTskH_CustIds%> </td>
      <td><%=vTskH_AccessIds%> </td>
      <td><%=fPassword(vTskH_Password)%>&nbsp; </td>
    </tr>
  <%
      '...any assets? 
      If vTskH_Level = 2 Then

        '...get assets
'       Response.Write "<br>TskH_No: " & vTskH_No
        sGetTskD_Rs vTskH_No      
        If Not oRs2.Eof Then
  %>
    <tr>
      <td colspan="10"style="text-align:center">        
        <div style="text-align:center" class="div" id="div<%=vTskH_No%>"><br>
        <table style="width:75%; margin:auto;">
          <tr>
            <th class="rowshade" style="width:10%">Type</th>
            <th class="rowshade" style="text-align:left; width:40%">&nbsp;Id</th>
            <th class="rowshade" style="text-align:left; width:40%">&nbsp;Title</th>
            <th class="rowshade" style="width:10%">Active</th>
          </tr>
  <%   
          Do While Not oRs2.Eof 
           sReadTskD                
          '...display title for program and module
          If Len(vTskD_Title) = 0 Then
            If Left(vTskD_Type, 1) = "M" Then 
              vTskD_Title = fModsTitle(vTskD_Id)
            ElseIf Left(vTskD_Type, 1) = "P" Then 
              vTskD_Title = fProgTitle(vTskD_Id)
            End If
          End If             
  %>
          <tr>
            <td style="text-align:center" width="10%">&nbsp;<%=vTskD_Type%> </td>
            <td style="width:10%">&nbsp;<%=fLeft(vTskD_Id, 32)%> </td>
            <td style="width:10%">&nbsp;<%=fLeft(vTskD_Title, 32) %> </td>
            <td style="text-align:center; width:10%">&nbsp;<%=fIf(vTskD_Active, "Y", "N")%></td>
          </tr>
  <% 
          oRs2.MoveNext
        Loop
        Set oRs2 = Nothing
        sCloseDb2 
  %>
        </table><br>
		  </div>		  
		  </td>
    </tr>    
  <%
      End If


      End If
      oRs.MoveNext
    Loop
    sCloseDB
  %>
    <tr>
      <td colspan="10"style="text-align:center">&nbsp;<p><a href="MyWorld.asp?vTskH_Id=<%=vTskH_Id%>">My Learning</a>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <a href="TaskEdit1.asp">Task Library</a><br>&nbsp;</p>
      </td>
    </tr>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


