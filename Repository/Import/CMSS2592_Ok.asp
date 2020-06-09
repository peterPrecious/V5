<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->

<%
  Dim oFs, oFile, vFile, vRecord, aRecord, vCnt, vError, aDept, vRecNo
  Dim v1, v2, v3, v4, v5, vb1, vb2, vb3, vb4, vb5, va1, va2, va3, va4, va5
  Dim aMemo, vMemo

  Server.ScriptTimeout = 60 * 10  '...allow 10 minutes for scripts

  '...keep open for entire run
  sOpenDb2 

  '...count users before update and store in "before variables"
  sCountUsers
  vb1 = v1:   vb2 = v2:   vb3 = v3:   vb4 = v4:   vb5 = v1+v2+v3+v4

  '...then inactivate all members
  sInactivateMembers

 '...imported names into a table
  Const ForReading = 1

  vMemb_AcctId  = Session("CustAcctId")
  svMembNo      = Session("MembNo")

  Set oFs = CreateObject("Scripting.FileSystemObject")   
  vFile = Server.MapPath("elearn.csv")
  Set oFile = oFs.OpenTextFile(vFile, ForReading)

  vRecNo = 0
  Do While oFile.AtEndOfStream <> True

    vRecNo = vRecNo + 1
'   If vRecNo Mod 1000  = 0 Then Stop

    vRecord = Replace(oFile.ReadLine, Chr(34), "")
    aRecord = Split(vRecord, ",")

    vMemb_Id          = fNoQuote(Ucase(aRecord(0)))
    vMemb_LastName    = fUnquote(aRecord(1))  
    vMemb_FirstName   = fUnquote(aRecord(2))  
    vMemb_Email       = fUnquote(aRecord(4))
    vMemb_Pwd         = Ucase(aRecord(5))
    vMemb_Memo        = Replace(fUnquote(aRecord(7)), "/", "|") & "|" & aRecord(8)
    vMemb_Active      = 1 
    vMemb_No          = spMembNoById (vMemb_AcctId, vMemb_Id, svMembNo)

    vSql = "UPDATE Memb SET"_
         & "  Memb_Id          = '" & vMemb_Id        & "', "_
         & "  Memb_LastName    = '" & vMemb_LastName  & "', "_  
         & "  Memb_FirstName   = '" & vMemb_FirstName & "', "_  
         & "  Memb_Email       = '" & vMemb_Email     & "', "_
         & "  Memb_Pwd         = '" & vMemb_Pwd       & "', "_
         & "  Memb_Memo        = '" & vMemb_Memo      & "', "_
         & "  Memb_Active      = 1  "_ 
         & "WHERE Memb_No = " & vMemb_No

    oDb2.Execute(vSql)
'   sTableUpdate "V5_Vubz", "Memb", vMemb_No

  Loop

  oFile.Close          

  '...count users after update and store in "after variables"
  sCountUsers
  va1 = v1:   va2 = v2:   va3 = v3:   va4 = v4:   va5 = v1+v2+v3+v4
  
  sCloseDb2

  '...inactive all active members (imported members will be active)
  Sub sInactivateMembers
    vSql = "UPDATE Memb SET Memb_Active = 0 WHERE Memb_AcctId = '2592' AND Memb_Level = 2 AND Memb_Internal = 0 AND Memb_Active = 1"
    oDb2.Execute(vSql)
  End Sub

  Sub sCountUsers
    vSql = "SELECT COUNT(*) AS Cnt FROM Memb WHERE Memb_AcctId = '2592' AND Memb_Level = 2 AND Memb_Internal = 0 And Memb_Active = 1"
    Set oRs = oDb2.Execute(vSql)
    v1 = oRs("Cnt")
    
    vSql = "SELECT COUNT(*) AS Cnt FROM Memb WHERE Memb_AcctId = '2592' AND Memb_Level = 2 AND Memb_Internal = 0 And Memb_Active = 0"
    Set oRs = oDb2.Execute(vSql)
    v2 = oRs("Cnt")

    vSql = "SELECT COUNT(*) AS Cnt FROM Memb WHERE Memb_AcctId = '2592' AND Memb_Level = 3 AND Memb_Internal = 0"
    Set oRs = oDb2.Execute(vSql)
    v3 = oRs("Cnt")

    vSql = "SELECT COUNT(*) AS Cnt FROM Memb WHERE Memb_AcctId = '2592' AND Memb_Level = 4 AND Memb_Internal = 0"
    Set oRs = oDb2.Execute(vSql)
    v4 = oRs("Cnt")

    Set oRs = Nothing
  End Sub

%>

<html>

<head>
  <title></title>
  <meta charset="UTF-8">
  <link href="<%=svDomain%>/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <link href="/V5/Inc/<%=Left(svCustId, 4)%>.css" type="text/css" rel="stylesheet">
  <% If vRightClickOff Then %><script language="JavaScript" src="/V5/Inc/RightClick.js"></script><% End If %>
  <script src="/V5/Inc/Functions.js"></script>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% 
    Server.Execute vShellHi 
  %>
  

  hello world 6




    <table border="0" cellpadding="3" style="border-collapse: collapse" bordercolor="#DDEEF9" width="100%">
    <tr>
      <td align="center">
      <h1>Upload (Import) Learner Profiles from City of Mississauga - Results</h1>
      <h2>Thank you.&nbsp; All learners have been uploaded.</h2>
      <p>&nbsp;</p>
      <table border="1" id="table2" cellspacing="0" cellpadding="10" style="border-collapse: collapse" bordercolor="#DDEEF9">
        <tr>
          <td class="c1" align="center">Totals</td>
          <td align="right" class="c1">Before</td>
          <td align="right" class="c1">After</td>
        </tr>
        <tr>
          <th align="right"># active learners : </th>
          <td align="right"><%=vb1%></td>
          <td align="right"><%=va1%></td>
        </tr>
        <tr>
          <th align="right"># inactive learners :</th>
          <td align="right"><%=vb2%></td>
          <td align="right"><%=va2%></td>
        </tr>
        <tr>
          <th align="right"># facilitators :</th>
          <td align="right"><%=vb3%></td>
          <td align="right"><%=va3%></td>
        </tr>
        <tr>
          <th align="right"># managers :</th>
          <td align="right"><%=vb4%></td>
          <td align="right"><%=va4%></td>
        </tr>
        <tr>
          <th align="right"># total: </th>
          <td align="right"><%=vb5%></td>
          <td align="right"><%=va5%></td>
        </tr>
      </table>
      <p>&nbsp;</p>
      </td>
    </tr>
  </table>
















  <!--#include virtual = "V5/Inc/Shell_LoLite.asp"-->
  
</body>

</html>