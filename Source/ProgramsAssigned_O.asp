<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<%
  Server.ScriptTimeout = 60 * 10
  Response.Buffer = False

  '...Determine the number of programs purchased
  Dim aProgs(), vCnt, aAssigned, vAssigned, vRow, custId, custAcctId

  custId = Request("custId")
  custAcctId = Right(custId, 4)

  vSql = " SELECT"_
       & "   Ecom.Ecom_Programs AS ProgId,"_
       & "   SUM(CASE WHEN Ecom_Amount < 0 THEN Ecom_Quantity * -1 ELSE Ecom_Quantity END) AS Purchased"_
       & " FROM Cust INNER JOIN Ecom ON Cust.Cust_AcctId = Ecom.Ecom_NewAcctId "_
       & " WHERE (Cust.Cust_AcctId = '" & custAcctId & "') AND Ecom_Archived IS NULL "_       
       & " GROUP BY Cust.Cust_Id, Ecom.Ecom_Programs "_
       & " ORDER BY Ecom.Ecom_Programs "

  vCnt = 0
  vRow = 0

  sOpenDb
  Set oRs = oDb.Execute(vSql)
  Do While Not oRS.eof
    vCnt = vCnt + 1
    ReDim Preserve aProgs (2, vCnt)
    aProgs (1, vCnt) = oRs("ProgId")
    aProgs (2, vCnt) = oRs("Purchased")
    oRs.MoveNext	  
  Loop
  Set oRs = Nothing
  sCloseDb
  

  If vCnt > 0 Then

    '...Determine the total number of programs assigned
    vSql = " SELECT Memb_Programs"_
         & " FROM Memb "_
         & " WHERE (Memb_AcctId = '" & custAcctId & "') "_
         & "   AND (LEN(Memb_Programs) > 6) "


'         & "   AND (Memb_Level IN (2, 3)) "_
'         & "   AND (Memb_Internal = 0) "

    sOpenDb
    Set oRs = oDb.Execute(vSql)
    Do While Not oRS.eof
      aAssigned = Split(oRs("Memb_Programs"))
      For i = 1 to vCnt
        For j = 0 To Ubound(aAssigned)
          If aProgs(1, i) = aAssigned(j) Then
            aProgs(0, i) = aProgs(0, i) + 1
            Exit For
          End If
        Next
      Next
      oRs.MoveNext	  
    Loop
    Set oRs = Nothing
    sCloseDb

  End If
  
%>

<html>

<head>
  <title>ProgramsAssigned_O</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>
    $(function() {
      $(".assign").click(function() { 
          $(".assign").css('background-color', '#FFFFFF');
          $(this).css('background-color', 'yellow');
          $(this).siblings().css('background-color', 'yellow');
        }
      );
    });
  </script>
  <style>
    .inactive { text-decoration: line-through; }
    .assign { border:1px solid navy; }
  </style>
</head>

<body>

  <% 
    Server.Execute vShellHi
  %>

  <table class="table" style="width: 350px;">

    <tr>
      <td colspan="<%=vCnt + 3%>">
        <p class="c1"><!--[[-->Programs Purchased and Assigned for<!--]]-->&nbsp;<%=custId %></p>
        <p class="c2"><!--[[-->This report displays all learners, sorted by Last Name, who have had programs assigned to them.<!--]]--></p>
        <p class="c3"><!--[[-->Summary...<!--]]--></p>
      </td>
    </tr>
    <tr>
      <td colspan="3" style="text-align: left;" class="rowshade"><!--[[-->Programs<!--]]--></td>
      <% For i = 1 To vCnt %>
      <td class="rowshade" style="text-align: center; width: 50px;"><a title="<%=fProgTitleClean(aProgs(1, i))%>" href="#"><%=aProgs(1, i)%></a></td>
      <% Next %>
    </tr>
    <tr>
      <td colspan="3" style="text-align: left;" class="rowshade"><!--[[-->Purchased Total<!--]]--></td>
      <% For i = 1 To vCnt %>
      <td class="rowshade" style="text-align: center;"><%=aProgs(2, i)%></td>
      <% Next %>
    </tr>
    <tr>
      <td colspan="3" style="text-align: left;" class="rowshade"><!--[[-->Assigned Total<!--]]--></td>
      <% For i = 1 To vCnt %>
      <td class="rowshade" style="text-align: center;"><%=aProgs(0, i)%></td>
      <% Next %>
    </tr>
    <tr>
      <td colspan="3" style="text-align: left;" class="rowshade"><!--[[-->Balance Remaining<!--]]--></td>
      <% For i = 1 To vCnt %>
      <td class="rowshade" style="text-align: center;"><%=aProgs(2, i) - aProgs(0, i)%></td>
      <% Next %>
    </tr>
    <tr>
      <td colspan="3" style="border-bottom:1px solid navy"><p class="c3"><!--[[-->Assigned...<!--]]--></p></td>
    </tr>

    <%
        Dim vActive
        vSql = " SELECT Memb_No, Memb_Id, Memb_FirstName, Memb_LastName, Memb_Programs, Memb_Active"_
             & " FROM Memb "_
             & " WHERE (Memb_AcctId = '" & custAcctId & "') "_
             & "   AND (Len(Memb_Programs) > 6) "_
             & " ORDER BY Memb_LastName, Memb_FirstName "


     '        & "   AND (Memb_Level IN (2, 3)) "_
     '        & "   AND (Memb_Internal = 0) "_


        sOpenDb
        Set oRs = oDb.Execute(vSql)
        Do While Not oRs.eof
          vActive = fIf(oRs("Memb_Active"), "", "inactive") 
        
    %>
    <tr>
      <td class="assign"><a href="User<%=fGroup%>.asp?vMembNo=<%=oRs("Memb_No")%>&vNext=ProgramsAssigned.asp"><%=oRs("Memb_Id")%></a></td>
      <td class="assign <%=vActive%>"><%=oRs("Memb_FirstName")%></td>
      <td class="assign <%=vActive%>"><%=oRs("Memb_LastName")%></td>
      <% 
          aAssigned = Split(oRs("Memb_Programs"))
          For i = 1 to vCnt
            For j = 0 To Ubound(aAssigned)
              If aProgs(1, i) = aAssigned(j) Then
                vAssigned = "<img border='0' src='../Images/Icons/Checkmark.gif' width='12' height='12'>"
                aProgs(0, i) = aProgs(0, i) + 1
                Exit For
              Else
                vAssigned = ""
              End If
            Next
      %>
      <td class="assign" style="text-align: center;"><%=vAssigned%></td>
      <% Next %>
    </tr>
    <%    
          oRs.MoveNext	  
          vRow = vRow + 1
          If vRow Mod 20 = 0 Then          
    %>

    <tr>
      <td colspan="<%=vCnt + 4%>">&nbsp;<br /></td>
    </tr>

    <tr>
      <th style="text-align: left" colspan="3" class="rowshade"><!--[[-->Programs<!--]]--></th>
      <% For i = 1 To vCnt %>
      <th class="rowshade"><a title="<%=fProgTitleClean(aProgs(1, i))%>" href="#"><%=aProgs(1, i)%></a></th>
      <% Next %>
    </tr>
    <%
          End If
        Loop
        Set oRs = Nothing
        sCloseDb
    %>
  </table>

  <br /><br />

  <input onclick="location.href = 'ProgramsAssigned.asp?custId=<%=custId%>'" type="button" value="<%=bReturn%>" name="bReturn" class="button070">

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>
</html>
