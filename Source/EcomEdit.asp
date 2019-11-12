<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->

<%

  Dim vSource, vLocked

  vSource = fDefault(Request("vSource"), "EcomReport0.asp")
  
  '...Get Ecom Transaction from Ecom Report or newly updated Ecom record or Delete request
  vEcom_No = Request.QueryString("vEcom_No")
  If Len(vEcom_No) > 0 Then

    sGetEcom
    If Request.QueryString("vAction") = "vDelete" Then '...delete
      sDeleteEcom
      Response.Redirect vSource
    End If

  '...PostBack for form
  ElseIf Request.Form("vHidden") = "Y" Then  
    sExtractEcom
    For Each vFld in Request.Form
      Select Case Left(vFld, 3)
        Case "vGo" '...go and get
          sGetEcom
          Exit For
        Case "vUp" '...update
          sUpdateEcom
          Exit For
        Case "vCl" '...insert
          vEcom_Issued = Now
          vEcom_Adjustment = True
          sInsertEcom
          Exit For
      End Select
    Next
  End If

  vLocked = fIf( (Month(Now) * Year(Now)) > (Month(vEcom_Issued) * Year(vEcom_Issued)), "true", "false")
  If svMembLevel = 5 Then vLocked = "false"


%>

<html>

<head>
  <title>EcomEdit</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <style>
    .note { background-color: yellow; border: 1px solid red; }
    th { padding-top: 8px; padding-right: 5px; }
    td { padding-top: 0px; }
  </style>
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <script>    
    function isPositive(id) { // check that quantity is a positive integer
      var val = id.value;
      var abs = Math.abs(val);
      var int = parseInt(val);
      if ((val != abs) || (val != int) || (int > 9999)) {
        alert("Quantity must be a positive integer less than 10000.");
        id.focus();
        return (false);
      }
      return (true);
    }

    $(function() {
      var locked = <%=vLocked%>;
      if (!locked){
        $(".note").hide();
      } else {
        alert ("Note: this record is closed.\nDollar and Date values cannot be modified \nnor can this record be deleted.");
        $(".note").show();
      }
    });      
  </script>

</head>

<body>

  <% Server.Execute vShellHi %>

  <% If svMembLevel = 5 Or svMembManager Then %>
  <h1>Clone &amp; Edit Ecommerce Logs</h1>
  <h2 class="c2">While ecommerce values (dollar and date) from previous months cannot be modified, nor can you delete previous month's records, you can modify or delete current month transactions.&nbsp; This page is best used to create new records (adjustments).&nbsp; To do this, access the appropriate&nbsp; transaction (via the Ecom No and <b>&quot;GO&quot;</b> below or via the Ecommerce Report link) then <b>Clone</b> a new record which will create a copy of the record shown with a new Ecommerce No.&nbsp; You can also add a record from scratch.</h2>
  <% Else %>
  <h1>Ecommerce Logs</h1>
  <h2>This allows you to view the details of an ecommerce transaction - you cannot update any of the fields.</h2>
  <% End If %>

  <form name="membs" method="POST" action="EcomEdit.asp" target="_self">
    <input class="c2" type="hidden" name="vHidden" value="Y">
    <input class="c2" type="hidden" name="vSource" value="<%=vSource%>">
    <table>
      <tr>
        <th>Ecommerce No :</th>
        <td>
          <input class="c2" type="text" size="8" name="vEcom_No" value="<%=vEcom_No%>">

          <% If svMembLevel = 5 Or svMembManager Then %>

          <!--            <input class="c2" border="0" src="../Images/Buttons/Go_<%=svLang%>.gif" name="vGo" type="image">-->
          <input type="submit" value="GO" name="vGo" class="button">

          <br />
          (Enter this number to retrieve a record. Once retrieved you cannot change the
            <br />
          value as this is system generated.&nbsp; If you add a new record a new value will be
            <br />
          assigned to the new record.)

          <% End If %>
        </td>
      </tr>
      <tr>
        <th>Customer Id :</th>
        <td>
          <input class="c2" type="text" size="15" name="vEcom_CustId" value="<%=vEcom_CustId%>"></td>
      </tr>
      <tr>
        <th>Account Id :</th>
        <td>
          <input class="c2" type="text" size="15" name="vEcom_AcctId" value="<%=vEcom_AcctId%>"></td>
      </tr>
      <tr>
        <th>Member Id :</th>
        <td>
          <input class="c2" type="text" size="46" name="vEcom_Id" value="<%=vEcom_Id%>"></td>
      </tr>
      <tr>
        <th>Member No :</th>
        <td>
          <input class="c2" type="text" size="10" name="vEcom_MembNo" value="<%=vEcom_MembNo%>"><br>
          Edit carefully, value must exist on member table</td>
      </tr>
      <tr>
        <th>Learner First Name :</th>
        <td>
          <input class="c2" type="text" size="26" name="vEcom_FirstName" value="<%=vEcom_FirstName%>"><br>
          This is the cardholder&#39;s first name, unless learner name entered</td>
      </tr>
      <tr>
        <th>Learner Last Name :</th>
        <td>
          <input class="c2" type="text" size="26" name="vEcom_LastName" value="<%=vEcom_LastName%>"><br>
          This is the cardholder&#39;s last name, unless learner name entered</td>
      </tr>
      <tr>
        <th>Cardholder Name :</th>
        <td>
          <input class="c2" type="text" size="46" name="vEcom_CardName" value="<%=vEcom_CardName%>"></td>
      </tr>
      <tr>
        <th>Address :</th>
        <td>
          <input class="c2" type="text" size="46" name="vEcom_Address" value="<%=vEcom_Address%>"></td>
      </tr>
      <tr>
        <th>City :</th>
        <td>
          <input class="c2" type="text" size="36" name="vEcom_City" value="<%=vEcom_City%>"></td>
      </tr>
      <tr>
        <th>Postal :</th>
        <td>
          <input class="c2" type="text" size="27" name="vEcom_Postal" value="<%=vEcom_Postal%>"></td>
      </tr>
      <tr>
        <th>Province :</th>
        <td>
          <input class="c2" type="text" size="27" name="vEcom_Province" value="<%=vEcom_Province%>"></td>
      </tr>
      <tr>
        <th>Country :</th>
        <td>
          <input class="c2" type="text" size="27" name="vEcom_Country" value="<%=vEcom_Country%>"></td>
      </tr>
      <tr>
        <th>Phone :</th>
        <td>
          <input class="c2" type="text" size="27" name="vEcom_Phone" value="<%=vEcom_Phone%>"></td>
      </tr>
      <tr>
        <th>Email :</th>
        <td>
          <input class="c2" type="text" size="46" name="vEcom_Email" value="<%=vEcom_Email%>"><br>
          Use the email address from Internet Secure,
          <br>
          which may differ from the member record.)</td>
      </tr>
      <tr>
        <th>Organization :</th>
        <td>
          <input class="c2" type="text" size="46" name="vEcom_Organization" value="<%=vEcom_Organization%>"></td>
      </tr>

      <tr>
        <th>Order Id :</th>
        <td>
          <input class="c2" type="text" size="24" name="vEcom_OrderId" value="<%=vEcom_OrderId%>">
          <br />
          Field added Jun 2018 - formerly in the shipping field</td>

      </tr>

      <tr>
        <th>Line Id :</th>
        <td>
          <input class="c2" type="text" size="24" name="vEcom_LineId" value="<%=vEcom_LineId%>">
          <br />
          Field added Jun 2018 - formerly in the shipping field</td>
      </tr>

      <tr>
        <th>Program Id :</th>
        <td>
          <input class="c2" type="text" size="10" name="vEcom_Programs" value="<%=vEcom_Programs%>"><br />
          Will be blank for SP learners.</td>
      </tr>

      <tr>
        <th>Catalogue No :</th>
        <td>
          <input class="c2" type="text" size="10" name="vEcom_CatlNo" value="<%=vEcom_CatlNo%>"><br />
          <!--<span style="background-color: yellow">New! Provides a Title for items.</span>-->
        </td>
      </tr>

      <tr>
        <th>Prices :</th>
        <td>
          <% If vLocked Then %>
          <input type="hidden" name="vEcom_Prices" value="<%=vEcom_Prices%>"><%=vEcom_Prices %>
          <% Else %>
          <input class="c2" type="text" size="10" name="vEcom_Prices" value="<%=vEcom_Prices%>">
          <% End If  %>
          <span class="note">Field is locked!</span><br>
          Program price before tax ie: 39. <font color="#FF0000">(See note on quantity field below.)</font></td>
      </tr>
      <tr>
        <th>Taxes :</th>
        <td>
          <% If vLocked Then %>
          <input type="hidden" name="vEcom_Taxes" value="<%=vEcom_Taxes%>"><%=vEcom_Taxes%>
          <% Else %>
          <input class="c2" type="text" size="10" name="vEcom_Taxes" value="<%=vEcom_Taxes%>">
          <% End If  %>
          <span class="note">Field is locked!</span><br>
          Total program taxes ie: 5.50.<font color="#FF0000"> (See note on quantity field below.)</font></td>
      </tr>
      <tr>
        <th>Issued :</th>
        <td>
          <% If vLocked Then %>
          <input type="hidden" size="10" name="vEcom_Issued" value="<%=fFormatSqlDate(vEcom_Issued)%>"><%=fFormatSqlDate(vEcom_Issued)%>
          <% Else %>
          <input class="c2" type="text" size="10" name="vEcom_Issued" value="<%=fFormatSqlDate(vEcom_Issued)%>">
          <% End If  %>
          <span class="note">Field is locked!</span><br>
          Date ordered via Internet Secure</td>
      </tr>
      <tr>
        <th>Expires :</th>
        <td>
          <input class="c2" type="text" size="10" name="vEcom_Expires" value="<%=fFormatSqlDate(vEcom_Expires)%>"><br>
          Expires date is the issue date plus the duration in days
          <br>
          from the customer program string - normally 90 days.<br />
          If you want to keep this transaction but do NOT want it
          <br />
          to appear in My Content, then set the Expiry Date either
          <br />
          to the Issue Date or any date before today.</td>
      </tr>
      <tr>
        <th>Amount :</th>
        <td>
          <% If vLocked Then %>
          <input type="hidden" size="10" name="vEcom_Amount" value="<%=vEcom_Amount%>"><%=vEcom_Amount%>
          <% Else %>
          <input class="c2" type="text" size="10" name="vEcom_Amount" value="<%=vEcom_Amount%>">
          <% End If  %>
          <span class="note">Field is locked!</span><br>
          <font color="#FF0000">(See note on quantity field below.)</font></td>
      </tr>
      <tr>
        <th>Currency :</th>
        <td>
          <input class="c2" type="text" size="26" name="vEcom_Currency" value="<%=vEcom_Currency%>"></td>
      </tr>
      <tr>
        <th>Quantity :</th>
        <td>
          <input class="c2" type="text" size="26" name="vEcom_Quantity" id="vEcom_Quantity" value="<%=vEcom_Quantity%>" onblur="isPositive(this)"><br>
          Can be zero for Group 1 License or 1+ seats.<br>
          <span class="red">Quantity must NEVER be negative.
          <!--<br />To remove a specific Quantity of programs, enter the Quantity as a positive number and enter negative amounts in the applicable dollar fields.-->
          </span></td>
      </tr>
      <tr>
        <th>New Account Id :</th>
        <td>
          <input class="c2" type="text" size="26" name="vEcom_NewAcctId" value="<%=vEcom_NewAcctId%>"><br>
          Automatically setup for group learning.</td>
      </tr>
      <tr>
        <th>Media :</th>
        <td>
          <input class="c2" type="text" size="26" name="vEcom_Media" value="<%=vEcom_Media%>"><br>
          Typically Online, Group2 or AddOn2.</td>
      </tr>
      <tr>
        <th>Shipping Label :</th>
        <td>
          <textarea rows="4" name="vEcom_Label" cols="38" class="c2"><%=vEcom_Label%></textarea><br>
          If empty, label is formed from address above.</td>
      </tr>
      <tr>
        <th>Order No :</th>
        <td>
          <input class="c2" type="text" size="26" name="vEcom_OrderNo" value="<%=vEcom_OrderNo%>"><br>
          Simple timestamp format: YYMM-DDhh-mmss</td>
      </tr>
      <tr>
        <th>Shipping Charge :</th>
        <td>
          <input class="c2" type="text" size="26" name="vEcom_Shipping" value="<%=vEcom_Shipping%>"></td>
      </tr>
      <tr>
        <th>Shipping Memo :</th>
        <td>
          <textarea rows="4" name="vEcom_Memo" cols="38" class="c2"><%=vEcom_Memo%></textarea><br>
          If empty, label is formed from address above.</td>
      </tr>
      <tr>
        <th>Payment Source :</th>
        <td>
          <input class="c2" type="radio" value="E" name="vEcom_Source" <%=fcheck("e", vecom_source)%>>
          Normal Ecommerce<br>
          <input class="c2" type="radio" value="V" name="vEcom_Source" <%=fcheck("v", vecom_source)%>>
          Manual to Vubiz<br>
          <input class="c2" type="radio" value="C" name="vEcom_Source" <%=fcheck("c", vecom_source)%>>
          Manual to Customer</td>
      </tr>
      <tr>
        <th>InternetSecure Receipt :</th>
        <td>
          <input class="c2" type="text" size="46" name="vEcom_InternetSecure" value="<%=vEcom_InternetSecure%>"></td>
      </tr>
      <tr>
        <th>Adjustment ?</th>
        <td>
          <input class="c2" type="radio" name="vEcom_Adjustment" value="0" <%=fcheck(fsqlboolean(vecom_adjustment), 0)%>>
          No (typically a live record)<br>
          <input class="c2" type="radio" name="vEcom_Adjustment" value="1" <%=fcheck(fsqlboolean(vecom_adjustment), 1)%>>
          Yes (typically a Cloned or Adjusted record)</td>
      </tr>
      <tr>
        <td colspan="2" style="text-align: center; padding-top: 50px;">

          <% If svMembLevel = 5 Or svMembManager Then %>
              To modify a transaction, first Clone the above record which you can then modify and Update.
              <span class="note">
                <br />
                <br />
                You cannot Delete nor Update dollar or dates from a previous month.</span><br />
          <br />
          <% End If %>

          <input onclick="history.back(1)" type="button" value="Return" name="bReturn" class="button070">

          <% If svMembLevel = 5 Or svMembManager Then %>

          <%=f10%>
          <input type="submit" value="Clone" name="vClone" class="button070"><%=f10%>

          <input type="submit" value="Update" name="vUpdate" class="button070"><%=f10%>
          <% If Not vLocked Then %>
          <input type="button" value="Delete" name="vDelete" class="button070" onclick="jconfirm('EcomEdit.asp?vAction=vDelete&vEcom_No=<%=vEcom_No%>&vSource=<%=vSource%>', 'Ok to delete?')"><%=f10%>
          <% End If %>
          <br />
          <br />
          <br />
          <a <%=fstatx%> href="EcomReport0.asp">Ecommerce Report</a>
          <% End If %>
          <br />
          <br />


        </td>
      </tr>
    </table>
  </form>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
