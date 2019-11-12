<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->

<%  	
  Dim vAddCustomerId, vEditCustId, vCloneCustId, vFunction, vCustId, vLeft, vAcctId, vParentId
  Dim vLevel, vActive, vCatMS, vOk, vNext, vTitle, vNextId, vCustPrefix, vCustAcctId

  vFunction   = Request("vFunction")
  vNext       = Request("vNext")


  '...this generates/inserts a new CustId then sends that to the Customer.asp where it will be specially "updated" 
  '   the trick is using cust_placeholder = true to mean that the customer record is written to lock up the code
  '   customer.asp will then update the record as though it was an insert which adds internals and repository
  If Request("vAddCustomerId").Count > 0 Then 
    vCustPrefix = Ucase(Request("vAddCustomerId"))
    vCustId = sp7nextCustId(vCustPrefix)
    Response.Redirect "Customer.asp?vAddCustId=" & vCustId 
  End If

  If Request("vCloneCustId").Count > 0 Then 
    vCustPrefix = Ucase(Request("vCloneCustId"))
    vCloneCustId = sp7nextCustId(vCustPrefix)
    Response.Redirect "Customer.asp?vCloneCustId=" & vCloneCustId & "&vCustId=" & Ucase(Request("vCustId")) 
  End If

  vLeft       = fDefault(Ucase(Trim(Request("vLeft"))), "")
  vAcctId     = fDefault(Request("vAcctId"), "")
  vParentId   = fDefault(Request("vParentId"), "")

  '...if we don't have a vLeft or vAcctId (from Customer.asp) create one
  vCustId     = fDefault(Request("vCustId"), "") 

  If Len(vCustId) >= 8 Then
    vLeft     = ""
    vAcctId   = ""
    vFunction = "list"
  End If

  vLevel      = fDefault(Request("vLevel"), "2")
  vActive     = Ucase(fDefault(Request("vActive"), "y"))
  vTitle      = Request("vTitle")
  vCatMs      = Ucase(fDefault(Request("vCatMs"), "n"))

  '...get next Customer Id (uses sp6nextAlpaha)
  Function sp7nextCustId(custPrefix)
    sp7nextCustId = ""
    sOpenCmdApp
    With oCmdApp
      .CommandText = "sp7nextCustId"     
      .Parameters.Append .CreateParameter("@custPrefix",  adVarChar, adParamInput,  4, custPrefix)
    End With
    Set oRsApp = oCmdApp.Execute()
    If Not oRsApp.Eof Then
      sp7nextCustId = oRsApp("nextCustId").Value
    End If
    Set oRsApp = Nothing
    Set oCmdApp = Nothing
    sCloseDbApp
  End Function 

%>

<html>

<head>
  <title>Customers</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>

    function emptyField(ele) {
      $(ele)[0].value = "";
    };

    function validate(ele) {
      var custCode = $(ele)[0].value;
      if (custCode.length != 4) {
        alert("Please enter a valid 4 character Customer prefix.");
        $(ele)[0].focus();
        return (false);
      } else if (custCode == "XXXX") {
        alert("Please replace the XXXX with a valid prefix.");
        $(ele)[0].focus();
        return (false);
      } else if (custCode.match(reAlpha)==null) { 
        alert("Customer Prefix must be ALPHA, ie CCHS.")
        $(ele)[0].focus();
        return (false);
      };
       return (true) 
    };

  </script>
</head>

<body>

  <% 
  	Server.Execute vShellHi
  %>

  <h1>Customers</h1>
  <table class="table">
    <tr>
      <!-- this is for the query parameters -->
      <td style="width: 65%; padding: 10px;">
        <p class="c2">Select starting Customers in one of three ways, max 50 then click <span class="code">Next</span>.&nbsp; From the list that should appear below, tap on the Customer Id to access the full Profile</p>
        <form method="POST" action="Customers.asp" target="_self">
          <input type="hidden" name="vFunction" value="list">
          <input type="hidden" value="<%=vNext%>" name="vNext">
          <table class="table">
            <tr>
              <th>Start with Customer Id : </th>
              <td>
                <input onfocus="emptyField('#vAcctId'); emptyField('#vParentId');" type="text" id="vCustId" size="9" value="<%=vCustId%>" name="vCustId">
                <a class="debug" onclick="emptyField('#vCustId')" href="#">&#937;</a>&nbsp; ie ABCD12345 or CCH</td>
            </tr>
            <tr>
              <th>or with Account Id : </th>
              <td>
                <input onfocus="emptyField('#vCustId'); emptyField('#vParentId');" type="text" id="vAcctId" size="5" value="<%=vAcctId%>" name="vAcctId">
                <a class="debug" onclick="emptyField('#vAcctId')" href="#">&#937;</a>&nbsp; ie 1234</td>
            </tr>
            <tr>
              <th>or with Parent Id : </th>
              <td>
                <input onfocus="emptyField('#vCustId'); emptyField('#vAcctId');" type="text" id="vParentId" size="5" value="<%=vParentId%>" name="vParentId">
                <a class="debug" onclick="emptyField('#vParentId')" href="#">&#937;</a>&nbsp; ie 1234</td>
            </tr>
            <tr>
              <th>whose Title contains : </th>
              <td>
                <input type="text" name="vTitle" size="32" value="<%=vTitle%>"></td>
            </tr>
            <tr>
              <th>that are : </th>
              <td style="background-color: yellow;">
                <input type="radio" value="2" name="vLevel" <%=fcheck(vlevel,  2)%> checked>Channel Parents or Children<br>
                <input type="radio" value="21" name="vLevel" <%=fcheck(vlevel, 21)%>>Channel Parents (Parent Id is empty)<br>
                <input type="radio" value="22" name="vLevel" <%=fcheck(vlevel, 22)%>>Channel Children (Parent Id points to Acct Id)<br>
                <input type="radio" value="4" name="vLevel" <%=fcheck(vlevel,  4)%>>Corporate<br>
                <input type="radio" value="0" name="vLevel" <%=fcheck(vlevel,  0)%>>Channel or Corporate              
              </td>
            </tr>
            <tr>
              <th>and is a : </th>
              <td>
                <input type="radio" value="Y" name="vCatMS" <%=fcheck(vcatMS, "y")%>>Catalogue Master or Sibling<br>
                <input type="radio" value="N" name="vCatMS" <%=fcheck(vcatMS, "n")%>>Catalogue Master or Sibling or NOT (ie doesn't affect search)</td>
            </tr>
            <tr>
              <th>and is : </th>
              <td>
                <input type="radio" value="Y" name="vActive" <%=fcheck(vactive, "y")%> checked>Active<br>
                <input type="radio" value="N" name="vActive" <%=fcheck(vactive, "n")%>>Inactive<br>
                <input type="radio" value="X" name="vActive" <%=fcheck(vactive, "x")%>>Active or Inactive
              </td>
            </tr>
            <tr>
              <td style="text-align: center" colspan="2">
                <br />
                <input type="submit" value="Next" name="bGo" class="button">
              </td>
            </tr>
          </table>
        </form>
      </td>

      <!-- this is for adding a new record -->
      <td style="text-align: center; width: 35%; padding: 10px;">
        <form method="POST" action="Customers.asp" onsubmit="return validate('#vAddCustomerId')">
          <p class="c2">To add a new Customer, replace the XXXX below with the 4 character Customer Prefix (ie CCHS) then tap <span class="code">Add</span>. The full Customer Id will appear on the next screen.</p>
          <input type="text" name="vAddCustomerId" id="vAddCustomerId" size="4" maxlength="4" value="XXXX">
          <a class="debug" onclick="emptyField('#vAddCustomerId')" href="#">&#937;</a>
          <input type="submit" value="Add" name="bAdd" class="button070">
        </form>
      </td>

    </tr>
  </table>

  <% If vFunction = "list" Then %>

  <table class="table">
    <tr>
      <th class="rowshade" style="width: 10%; text-align: center;">Customer Id</th>
      <th class="rowshade" style="width: 05%; text-align: center;">Acct Id</th>
      <th class="rowshade" style="width: 05%; text-align: center;">Parent Id</th>
      <th class="rowshade" style="width: 05%; text-align: center;">Type</th>
      <th class="rowshade" style="width: 40%; text-align: left;">Title</th>
      <th class="rowshade" style="width: 20%; text-align: center; white-space: nowrap;">Clone into<br />Customer Id staring with...</th>
    </tr>
    <%
      '...read Cust
      vSql  = " SELECT TOP 50 * FROM Cust WHERE " _
            & " Cust_Title LIKE '%" & vTitle & "%'" _
            & fIf(Len(vCustId) > 0, " AND Cust_Id > = '" & vCustId & "'", "") _
            & fIf(Len(vAcctId) > 0, " AND Cust_AcctId = '" & vAcctId & "'", "") _
            & fIf(Len(vParentId) > 0, " AND Cust_ParentId = '" & vParentId & "'", "") _

            & fIf(vLevel = 2, " AND Cust_Level = 2", "") _
            & fIf(vLevel = 4, " AND Cust_Level = 4", "") _
            & fIf(vLevel = 21, " AND Cust_Level = 2 AND LEN(Cust_ParentId) = 0", "" ) _
            & fIf(vLevel = 22, " AND Cust_Level = 2 AND LEN(Cust_ParentId) = 4", "" ) _

            & fIf(vCatMS = "Y", " AND (Cust_CatalogueMaster = 1 OR Cust_CatalogueSibling = 1)", "") _
      
            & fIf(vActive = "Y", " AND Cust_Active = 1", "") _
            & fIf(vActive = "N", " AND Cust_Active = 0", "") _
            & fIf(Len(vAcctId) > 0, " ORDER BY Cust_AcctId, Cust_Id ", " ORDER BY Cust_Id ") 
 
'     sDebug
      sOpenDb
      Set oRs  = oDb.Execute(vSql)
      i = 0
      Do While Not oRs.Eof
        sReadCust
        i = i + 1  
        vOk = True
        If vOk Then
    %>
    <tr>
      <td style="text-align: center">
        <input onclick="location.href = 'Customer.asp?vEditCustId=<%=vCust_Id%>'" type="button" value="<%=vCust_Id%>" name="bAdd" class="button100"></td>
      <td style="text-align: center"><%=vCust_AcctId%></td>
      <td style="text-align: center"><%=vCust_ParentId%></td>
      <td style="text-align: center"><%=fIf(vCust_Level = "2", "Channel", "Corporate")%></td>
      <td><%=fLeft(vCust_Title, 40)%></td>
      <td style="text-align: center">
        <form method="POST" action="Customers.asp" onsubmit="return validate('.clone_<%=i%>')">
          <input type="text" id="vCloneCustId" name="vCloneCustId" class="clone_<%=i%>" size="4" maxlength="4" value="<%=Left(vCust_Id, 4)%>">
          <a class="debug" onclick="emptyField('.clone_<%=i%>')" href="#">&#937;</a>
          <input type="submit" value="Clone" name="bClone" class="button">
          <input type="hidden" name="vCustId" value="<%=vCust_Id%>">
        </form>
      </td>
    </tr>
    <%  
        End If
        oRs.MoveNext
      Loop
      Set oRs = Nothing
      sCloseDB    
    %>

    <script>
      <% If Len(vParentId) > 0 Then %>
        $("#vCustId")[0].value = "";
        $("#vAcctId")[0].value = "";
        $("#vParentId")[0].value = "<%=vCust_ParentId%>";
      <% ElseIf Len(vAcctId) > 0 Then %>
        $("#vCustId")[0].value = "";
        $("#vAcctId")[0].value = "<%=vCust_AcctId%>";
        $("#vParentId")[0].value = "";
      <% Else %>
        $("#vCustId")[0].value = "<%=vCust_Id%>";
        $("#vAcctId")[0].value = "";
        $("#vParentId")[0].value = "";
      <% End If %>
    </script>

    <tr>
      <td colspan="6" style="text-align: center">
        <input onclick="location.href = 'Customers.asp?vLeft=<%=vLeft%>&vAcctId=<%=Right(vCust_Id, 4)%>&vParentId=<%=Right(vParentId,4)%>&vFunction=list'" type="button" value="Next" name="bNext" class="button085">
      </td>
    </tr>

  </table>

  <% End If %>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
