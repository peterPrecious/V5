<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Routines.asp"-->

<%
  Dim vCnt, vTaxType, vTaxRate, vLearners, vSource, vCurLearners, vCustExpires, vPrice, vAmount, vTax, vTotal, vTotAmount, vTotTax, vTotTotal, vEcomNo, vDate, vUrl, vTotLearners
  
  '...Get Ecom Transaction from Ecom Report or newly updated Ecom record or Delete request
  vEcom_NewAcctId = Request("vEcom_NewAcctId")
  vEcom_CustId    = fEcomCustId (vEcom_NewAcctId)
  If Len(vEcom_CustId) <> 8 Then Response.Redirect "Menu.asp"

  vLearners = fDefault(Request("vLearners"), 5)
  vSource   = fDefault(Request("vSource"), "C")
  
  '...PostBack from form
  If Request("bUpdate").Count > 0 Then  
    '...add in the ecom adjustment records
    For Each vFld In Request.Form
      If Left(vFld, 7)   = "vPrice_" Then
        vEcomNo          = Mid(vFld, 8)
        vEcom_No         = Clng(vEcomNo)
        sGetEcom
        vEcom_Prices     = Request(vFld) * vLearners
        vEcom_Taxes      = Request("vTaxes_"& vEcomNo)
        vEcom_Amount     = Request("vAmount_"& vEcomNo)
        vEcom_Quantity   = vLearners
        vEcom_Issued     = Now
        vEcom_Source     = vSource
        vEcom_Adjustment = True
        sInsertEcom
      End If
    Next

    '...update the Max Learners field in customer record
    sUpdateCustMaxUsers Left(vEcom_CustId, 4) & vEcom_NewAcctId, vLearners
  End If

  vCurLearners = fCustMaxUsers(Left(vEcom_CustId, 4) & vEcom_NewAcctId)
  vCustExpires = fCustG1Expires(Left(vEcom_CustId, 4) & vEcom_NewAcctId)
  
%>


<html>

<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

</head>

<body>

<% 
  Server.Execute vShellHi 
%>

  <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vLearners.value == "")
  {
    alert("Please enter a value for the \"New Learners\" field.");
    theForm.vLearners.focus();
    return (false);
  }

  if (theForm.vLearners.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"New Learners\" field.");
    theForm.vLearners.focus();
    return (false);
  }

  if (theForm.vLearners.value.length > 3)
  {
    alert("Please enter at most 3 characters in the \"New Learners\" field.");
    theForm.vLearners.focus();
    return (false);
  }

  var checkOK = "0123456789-";
  var checkStr = theForm.vLearners.value;
  var allValid = true;
  var validGroups = true;
  var decPoints = 0;
  var allNum = "";
  for (i = 0;  i < checkStr.length;  i++)
  {
    ch = checkStr.charAt(i);
    for (j = 0;  j < checkOK.length;  j++)
      if (ch == checkOK.charAt(j))
        break;
    if (j == checkOK.length)
    {
      allValid = false;
      break;
    }
    allNum += ch;
  }
  if (!allValid)
  {
    alert("Please enter only digit characters in the \"New Learners\" field.");
    theForm.vLearners.focus();
    return (false);
  }

  var chkVal = allNum;
  var prsVal = parseInt(allNum);
  if (chkVal != "" && !(prsVal >= 1 && prsVal <= 999))
  {
    alert("Please enter a value greater than or equal to \"1\" and less than or equal to \"999\" in the \"New Learners\" field.");
    theForm.vLearners.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form name="FrontPage_Form1" method="POST" action="EcomAdjustG1.asp" target="_self" onsubmit="return FrontPage_Form1_Validator(this)" language="JavaScript">

    <input type="hidden" name="vHidden" value="Y">
    <input type="hidden" name="vEcom_NewAcctId" value="<%=vEcom_NewAcctId%>">
    <table border="0" width="100%" cellpadding="10" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td width="100%" align="center">
        <% If svMembLevel = 5 Or svMembManager Then %>
        <h1 align="center">Adjust Ecommerce Transactions for Group 1 Sales</h1><h2 align="left">This allows you increase (or decrease) the number of learners/programs in a Group 1 Site by adding an <b>Adjustment</b> record for each program originally ordered.&nbsp; <span style="background-color: #FFFF00">Enter fields highlighted in yellow</span>&nbsp;then click <b>Calculate</b> to confirm total cost of the transaction. When ready click <b>Update</b> making the adjustments permanent.</h2>
        <% Else %>
        <h1 align="center">Ecommerce Transactions for Group 1 Sales</h1><h2>This shows the original ecommerce transactions and any subsequent adjustments.</h2>
        <% End If %>
        </td>
      </tr>
      <tr>
        <td width="100%" align="center">
        <table border="1" id="table4" cellspacing="0" cellpadding="5" style="border-collapse: collapse" bordercolor="#DDEEF9" width="600">
          <tr>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" colspan="6">Original Sale</th>
          </tr>
          <tr>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF">New Acct</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF">Channel</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF">Cardholder</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF">Issued</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF">Expires</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF">Currency</th>
          </tr>

<%
  vCnt         = 0
  vTotAmount   = 0
  vTotTax      = 0
  vTotTotal    = 0
  vTotLearners = 0

  sGetEcomByNewAcctId_Rs (vEcom_NewAcctId)
  Do While Not oRs.Eof
    sReadEcom   
  
    '...display basic info from first record
    If vCnt = 0 Then

      vDate = vEcom_Issued '...this ensures we are dealing with an original not an adjustment
      '...determine tax
      vTaxRate = fGST(Now, vEcom_Country, vEcom_Province) 
      If vTaxRate > 0 Then 
        vTaxType = "GST" 
      Else  
        vTaxRate = fHST(Now, vEcom_Country, vEcom_Province) 
        If vTaxRate > 0 Then 
          vTaxType = "HST" 
        End If
      End If
%>

          <tr>
            <td align="center"><%=Left(vEcom_CustId, 4) & vEcom_NewAcctId%></td>
            <td align="center"><%=vEcom_CustId%></td>
            <td align="center"><%=vEcom_FirstName & " " & vEcom_LastName%></td>
            <td align="center"><%=fFormatSqlDate(vEcom_Issued)%></td>
            <td align="center"><%=fFormatSqlDate(vEcom_Expires)%></td>
            <td align="center"><%=vEcom_Currency%></td>
          </tr>
        </table>
        </td>
      </tr>
      <tr>
        <td width="100%" align="center">
        <table border="1" id="table5" cellspacing="0" cellpadding="5" style="border-collapse: collapse" bordercolor="#DDEEF9">
          <tr>
            <th align="right">Original # of Learners Purchased :</th>
            <td align="right"><%=vEcom_Quantity%></td>
          </tr>
          <tr>
            <th align="right">Current # of Learners in Customer Profile :</th>
            <td align="right"><%=vCurLearners%></td>
          </tr>
          <tr>
            <th align="right">Current Expiry Date in Customer Profile :</th>
            <td align="right"><%=fFormatDate(vCustExpires)%></td>
          </tr>
          </table>
        </td>
      </tr>
      <tr>
        <td width="100%" align="center" height="266">
        <table border="1" id="table6" cellspacing="0" cellpadding="2" style="border-collapse: collapse" bordercolor="#DDEEF9" width="600">
          <tr>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" colspan="9" height="30">Transaction History</th>
          </tr>
          <tr>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Issued</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Programs</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Type</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Source</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30"># of<br>Learners</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" align="right">Price<br>Each</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" align="right">Total<br>pre Tax</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" align="right"><%=fIf(vTaxType <> "", vTaxType & "<br>" & vTaxRate * 100 & "%", "")%></th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" align="right">Total<br>with Tax</th>
          </tr>

<%
    End If       
    
    '...only pickup the 2nd record (#seats) as the first is the license   
    If Not vEcom_Adjustment Then vCnt = vCnt + 1
    
    If vCnt Mod 2 = 0 Then
    
      '...ignore free courses
      If vEcom_Prices > 0 Then
        vUrl = "<a href='EcomEdit.asp?vEcom_No=" & vEcom_no & "&vSource=" & Server.UrlEncode("EcomAdjustG1.asp?vEcom_NewAcctId=" & vEcom_NewAcctId) & "'>" & vEcom_Programs & "</a>"

        '...display total number of learners    
        vTotLearners = vTotLearners + vEcom_Quantity

        '...display new adjustments in greeen
        If fFormatDate(vEcom_Issued) = fFormatDate(Now) Then
%>
            <tr>
              <td align="center" height="30"><font color="#008000"><%=fFormatDate(vEcom_Issued)%></font></td>
              <td align="center" height="30"><font color="#008000"><%=vUrl%></font></td>
              <td align="center" height="30"><font color="#008000"><b><%=fIf(vEcom_Adjustment, "Adj", "Orig")%></b></font></td>
              <td align="center" height="30"><font color="#008000"><%=vEcom_Source%></font></td>
              <td align="center" height="30"><font color="#008000"><%=vEcom_Quantity%></font></td>
              <td align="right" height="30"><font color="#008000"><%=FormatCurrency(vEcom_Prices/vEcom_Quantity, 2)%></font></td>
              <td align="right" height="30"><font color="#008000"><%=FormatCurrency(vEcom_Prices, 2)%></font></td>
              <td align="right" height="30"><font color="#008000"><%=FormatCurrency(vEcom_Taxes, 2)%></font></td>
              <td align="right" height="30"><font color="#008000"><%=FormatCurrency(vEcom_Amount)%></font></td>
          </tr>
<%
        Else
%>                    
            <tr>
              <td align="center" height="30"><%=fFormatDate(vEcom_Issued)%></td>
              <td align="center" height="30"><%=vUrl%></td>
              <td align="center" height="30"><b><%=fIf(vEcom_Adjustment, "Adj", "Orig")%></b></td>
              <td align="center" height="30"><b><%=vEcom_Source%></b></td>
              <td align="center" height="30"><%=vEcom_Quantity%></td>
              <td align="right" height="30"><%=FormatCurrency(vEcom_Prices/vEcom_Quantity, 2)%></td>
              <td align="right" height="30"><%=FormatCurrency(vEcom_Prices, 2)%></td>
              <td align="right" height="30"><%=FormatCurrency(vEcom_Taxes, 2)%></td>
              <td align="right" height="30"><%=FormatCurrency(vEcom_Amount)%></td>
          </tr>
<%
        End If
      End If
    End If
    oRs.MoveNext
  Loop
%>
            <tr>
              <td align="center" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">&nbsp;</td>
              <td align="center" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">&nbsp;</td>
              <td align="center" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">&nbsp;</td>
              <td align="center" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">&nbsp;</td>
              <td align="center" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF"><b><%=vTotLearners%></b></td>
              <td align="right" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">&nbsp;</td>
              <td align="right" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">&nbsp;</td>
              <td align="right" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">&nbsp;</td>
              <td align="right" height="30" bgcolor="#DDEEF9" bordercolor="#FFFFFF">&nbsp;</td>
          </tr>

<% If svMembLevel = 5 Or svMembManager Then %>

          <tr>
            <th bordercolor="#FFFFFF" height="30" colspan="9">&nbsp;
              <table border="1" id="table7" cellspacing="0" cellpadding="4" bordercolor="#DDEEF9" style="border-collapse: collapse">
                <tr>
                  <th align="right"># of Learners to Add :</th>
                  <th align="right" bgcolor="#FFFF00">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <!--webbot bot="Validation" s-display-name="New Learners" s-data-type="Integer" s-number-separators="x" b-value-required="TRUE" i-minimum-length="1" i-maximum-length="3" s-validation-constraint="Greater than or equal to" s-validation-value="1" s-validation-constraint="Less than or equal to" s-validation-value="999" -->
                    <input type="text" name="vLearners" size="2" maxlength="3" style="text-align: right" value="<%=vLearners%>">&nbsp;
                  </th>
                </tr>
                <tr>
                  <td align="center" colspan="2"><b>Transaction Source</b>:&nbsp;<br>E: Normal Ecommerce<br>V: Manual Payment to Vubiz<br>C: Manual Payment to Channel<br>&nbsp;</td>
                </tr>
                </table>&nbsp;
            </th>
          </tr>

          <tr>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Issued</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Programs</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Type</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Source</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30"># of<br>Learners</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Price<br>Each</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" align="right">Total<br>pre Tax</th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" align="right"><%=fIf(vTaxType <> "", vTaxType & "<br>" & vTaxRate * 100 & "%", "")%></th>
            <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" align="right">Total<br>with Tax</th>
          </tr>

<% 
  vCnt = 0
  sGetEcomByNewAcctId_Rs (vEcom_NewAcctId)
  Do While Not oRs.Eof
    sReadEcom   
   
    '...only pickup the 2nd record (#seats) as the first is the license   
    If Not vEcom_Adjustment Then 
      
      vCnt = vCnt + 1
      
      If vCnt Mod 2 = 0 Then
      
       '...ignore free courses
       If vEcom_Prices > 0 Then
         vPrice = fDefault(Request("vPrice_" & vEcom_No), vEcom_Prices/vEcom_Quantity)
      
         vAmount = vPrice * vLearners
         vTax    = vPrice * vLearners * vTaxRate
         vTotal  = vAmount + vTax 
                            
         If Not vEcom_Adjustment Then  
           vTotAmount = vTotAmount + vAmount
           vTotTax    = vTotTax    + vTax
           vTotTotal  = vTotTotal  + vTotal
         End If
         
%>
            <tr>
              <td align="center" height="30"><%=fFormatDate(Now)%></td>
              <td align="center" height="30"><%=vEcom_Programs%></td>
              <td align="center" height="30"><b>New</b></td>
              <td align="center" height="30" bgcolor="#FFFF00">
                <select size="1" name="vSource">
                  <option value="E" <%=fSelect(vSource, "E")%>>E</option>
                  <option value="C" <%=fSelect(vSource, "C")%>>C</option>
                  <option value="V" <%=fSelect(vSource, "V")%>>V</option>
                </select>
              </td>
              <td align="center" height="30"><%=vLearners%></td>
              <td align="center" height="30" bgcolor="#FFFF00"><input type="text" name="vPrice_<%=vEcom_No%>" size="6" value="<%=FormatNumber(vPrice, 2)%>" style="text-align: right; font-family:Verdana; font-size:8pt">&nbsp;</td>
              <td align="right" height="30"><%=FormatCurrency(vAmount, 2)%></td>
              <td align="right" height="30"><%=FormatCurrency(vTax, 2)%></td>
              <td align="right" height="30"><%=FormatCurrency(vTotal, 2)%></td>
            </tr>

            <input type="hidden" name="vTaxes_<%=vEcom_No%>" value="<%=vTax%>">
            <input type="hidden" name="vAmount_<%=vEcom_No%>" value="<%=vTotal%>">


<%
        End If
      End If
    End If
    oRs.MoveNext
  Loop
%>
            <tr>
              <th colspan="6" bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" align="right">Total New Adjustments $&nbsp;&nbsp;&nbsp; </th>
              <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" align="right"><%=FormatCurrency(vTotAmount, 2)%></th>
              <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" align="right"><%=FormatCurrency(vTotTax, 2)%></th>
              <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30" align="right"><%=FormatCurrency(vTotTotal, 2)%></th>
          </tr>

<% End If %>


            <tr>
              <td colspan="9" align="center">
             
                <% If svMembLevel = 5 Or svMembManager Then %>
  
                <h2><br>Before you <b>Update</b>, generate Totals by clicking <b>Calculate</b>.</h2>
                <p><input type="submit" value="Calculate" name="bCalculate" class="button"></p>
                <h2 align="left">If the above values are correct,&nbsp; ie you have not changed any &quot;yellow&quot; fields since you clicked <b>Calculate</b>, then click <b>Update</b> and your <b>New</b> transactions will become permanent <b>Adj</b>ustments above.&nbsp; Adjustments made today will appear in green.</h2>
                <p><input type="submit" value="Update" name="bUpdate" class="button"></p><h2 align="center">
                <a <%=fstatx%> href="EcomReport0.asp">Ecommerce Report</a></h2>
        
                <% Else %>
  
                <p><input type="button" onclick="location.href='javascript:history.back(1)'" value="Return" name="bReturm" class="button"></td>
          
  			        <% End If %>          

              </td>
          </tr>


        </table>





      </tr>

    </table>
    <input type="hidden" name="vTotLearners" value="<%=vTotLearners%>">
  </form>

<% 
  Server.Execute vShellLo 
%>

</body>

</html>