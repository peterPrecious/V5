<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Discounts.asp"-->

<%
  Dim aValues, aValue, vTotal, vOnFile, vCnt, vTitle, vBg, vOk, vTotUS, vTotCA, vTotQty, vStr, vEnd, vOnLoadScript

  vEcom_Media = Session("Ecom_Media")
  If fNoValue(vEcom_Media) Then Response.Redirect "EcomError.asp?vMsg=" & Server.UrlEncode("No media selected in " & Request.ServerVariables("Script_Name"))

  '...this is used to refresh the CD programs page when it selects a CD from the left panel, not used for Online, Group or Prords as they are "underneath" each other
  If vEcom_Media = "CDs" Then
    vOnLoadScript = "onLoad=" & chr(34) 
    vOnLoadScript = vOnLoadScript & "parent.frames.Left.location.href='Ecom2ProgramsCDs.asp';" 
    vOnLoadScript = vOnLoadScript & chr(34)
  Else
    vOnLoadScript = ""
  End If    

  '...determine if any discounts apply from the customer file
  sGetCust svCustId

  '...keep track which form called so can return
  vPage = Request("vPage")

  '...retrieve/create the basket info
  svProdNo       = Session("ProdNo")
  svProdMax      = Session("ProdMax")
  If svProdNo > 0 Then  
    saProd       = Session("Prod")
  Else
    Dim saProd()
    ReDim Preserve saProd(6, svProdNo) '... note: discount % held in 0
  End If

  '...see if clearing basket
  If Request("vAction") = "Clear" Then
    svProdNo  = 0
    svProdMax = 0
    ReDim Preserve saProd(6, svProdNo)
  End If

  '...see if updating Group quantity 
  If Len(Request.Form("vEcom_Quantity")) > 0 Then
    vEcom_Quantity  = Request("vEcom_Quantity")
    Session("Ecom_Quantity") = vEcom_Quantity

    '...update product array with new quantity
    For i = 1 to svProdMax
      If saProd(2, i) > 0 Then 
        '...don't update number of group licenses
        If Not (vEcom_Media = "Group" And Instr(saProd(6, i), "Annual License:") > 0) Then
          saProd(2, i) = vEcom_Quantity
        End If
        '...update seat descriptions
        If vEcom_Media = "Group" And Instr(saProd(6, i), "Seats:") > 0 Then
          saProd(6, i) = "<b>" & vEcom_Quantity & Mid(saProd(6, i), Instr(saProd(6, i), " Seats:"))
        End If
      End If
    Next 
    Session("Prod") = saProd       
  '...else get from session variable
  Else
    vEcom_Quantity = Session("Ecom_Quantity")
  End If


  '...see if updating CD quantity
  For Each vFld In Request.Form
    If Left(vFld, 5) = "vQty_" Then
      i = Cint(Mid(vFld, 6))
      saProd(2, i) = Request(vFld)      
      '...update CD descriptions
      If Instr(saProd(6, i), "CDs:") = 0 Then
        If saProd(2, i) > 1 Then
          saProd(6, i) = "<b>" & saProd(2, i) & " CDs: </b>" & saProd(6, i)
        End If
      Else
        If saProd(2, i) = 1 Then
          saProd(6, i) = Mid(saProd(6, i), Instr(saProd(6, i), "CDs:") + 9)
        Else
          saProd(6, i) = "<b>" & saProd(2, i) & Mid(saProd(6, i), Instr(saProd(6, i), "CDs:"))
        End If
      End If      
    End If
  Next
  

  If Len(Request.Form("vEcom_Quantity")) > 0 Then
    vEcom_Quantity  = Request("vEcom_Quantity")
    Session("Ecom_Quantity") = vEcom_Quantity
    '...update product array with new quantity
    For i = 1 to svProdMax
      If saProd(2, i) > 0 Then 
        '...don't update number of group licenses
        If Not (vEcom_Media = "Group" And Instr(saProd(6, i), "Annual License:") > 0) Then
          saProd(2, i) = vEcom_Quantity
        End If
        '...update seat descriptions
        If vEcom_Media = "Group" And Instr(saProd(6, i), "Seats:") > 0 Then
          saProd(6, i) = "<b>" & vEcom_Quantity & Mid(saProd(6, i), Instr(saProd(6, i), " Seats:"))
        End If
      End If
    Next 
    Session("Prod") = saProd       
  '...else get from session variable
  Else
    vEcom_Quantity = Session("Ecom_Quantity")
  End If

  '...build quantity title, leave empty for Online
  

  If vEcom_Media = "Group" Then
    vTitle = fPhraH(000234)
  ElseIf vEcom_Media = "CDs" Then
    vTitle = fPhraH(000075)
  Else
    vTitle = ""
  End If


  '...add the selected item into the array
  If Request("vProgram").Count > 0 Then

    vValue  = Replace(Request("vProgram"),"'"," ")
    aValues = Split (vValue,"||")

    For i = 0 To Ubound(aValues)
      aValue  = Split (aValues(i),"~")
      '...remove unnecessary text from program name
      j = Instr(Lcase(aValue(6)), "<br>")
      If j > 0 Then aValue(6) = Left(aValue(6), j-1)

      
      '...ensure this product is not on file
      vOnFile = False
      If svProdMax > 0 Then
        For j = 1 to svProdMax
          If saProd(1, j) = aValue(0) Then 
            svProdNo = j + 1 '...add 1 to the location as there are always pairs of products
            vOnFile = True
            Exit For
          End If
        Next 
      End If
  
      '...if not on file, then add item to bottom of array
      If Not vOnFile Then
  
        '...if Group, add an extra entry for the license
        If vEcom_Media = "Group" Then
          svProdNo             = svProdNo + 1
          svProdMax            = svProdMax + 1
          ReDim Preserve saProd(6, svProdNo)
    
          saProd(0, svProdNo) = 0                                          '...percentage discount
          saProd(1, svProdNo) = aValue(0) 							                   '...program Id
          saProd(2, svProdNo) = 1                                          '...quantity always 1 license per annum
          saProd(3, svProdNo) = aValue(3)                                  '...US price lic
          saProd(4, svProdNo) = aValue(4)                                  '...CA price lic
          saProd(5, svProdNo) = 365                                        '...license is always a duration of 365 days
          saProd(6, svProdNo) = "<b>Annual License:</b> " & aValue(6)      '...description
    
          '...get totals for discount calcs
          vTotUS  = vTotUS  + saProd(3, svProdNo) * saProd(2, svProdNo)
          vTotCA  = vTotCA  + saProd(4, svProdNo) * saProd(2, svProdNo)
          vTotQty = vTotQty + saProd(2, svProdNo)
        End If
    
        '...now add the item 
        svProdNo             = svProdNo + 1
        svProdMax            = svProdMax + 1
        ReDim Preserve saProd(6, svProdNo)
      
        saProd(0, svProdNo) = 0                              '...percentage discount
        saProd(1, svProdNo) = aValue(0)                      '...program Id
        saProd(2, svProdNo) = vEcom_Quantity                 '...quantity
        saProd(3, svProdNo) = aValue(1)                      '...US price each
        saProd(4, svProdNo) = aValue(2)                      '...CA price each
        saProd(5, svProdNo) = aValue(5)                      '...no days duration
        saProd(6, svProdNo) = fIf(vEcom_Media = "Group", "<b>" & vEcom_Quantity & " Seats:</b> ", "") & aValue(6)      '...preface group sales description
  
        '...get totals for discount calcs
        vTotUS  = vTotUS  + saProd(3, svProdNo) * saProd(2, svProdNo)
        vTotCA  = vTotCA  + saProd(4, svProdNo) * saProd(2, svProdNo)
        vTotQty = vTotQty + saProd(2, svProdNo)
  
        '...save basket values
        Session("ProdNo")    = svProdNo
        Session("ProdMax")   = svProdMax
        Session("Prod")      = saProd       
  
      End If

   Next

  End If

  '...put any discounts into cell 0  
  If vEcom_Media <> "Group" Then
    sCheckDiscounts
  End If
  
  '...save basket values
  Session("ProdNo")  = svProdNo
  Session("ProdMax") = svProdMax
  Session("Prod")    = saProd       
%>

<html>

<head>
  <title>Ecom2Basket</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>AC_FL_RunContent = 0;</script>
  <script src="/V5/Inc/AC_RunActiveContent.js" language="javascript"></script>
  <script>
    function jTitle (vTitle, vImage) {
      var vParm = "title=" + vTitle + '&image=/V5/Images/Titles/' + vImage;
      AC_FL_RunContent('codebase','//download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0','name','flashVars','width','265','height','85','align','middle','id','flashVars','src','/V5/Images/Titles/VuTitles','FlashVars',vParm,'quality','high','bgcolor','#ffffff','allowscriptaccess','sameDomain','allowfullscreen','false','pluginspage','///go/getflashplayer','movie','/V5/Images/Titles/VuTitles');
    }

    function submitForm(theForm)
    {
      theForm.action = 'Ecom2Basket.asp';
      theForm.submit();
    }

    function FrontPage_Form1_Validator(theForm) {

      if (theForm.vEcom_Quantity.value == "") {
        alert("Please enter a value for the \"No Seats\" field.");
        theForm.vEcom_Quantity.focus();
        return (false);
      }

      if (theForm.vEcom_Quantity.value.length < 1) {
        alert("Please enter at least 1 characters in the \"No Seats\" field.");
        theForm.vEcom_Quantity.focus();
        return (false);
      }

      if (theForm.vEcom_Quantity.value.length > 3) {
        alert("Please enter at most 3 characters in the \"No Seats\" field.");
        theForm.vEcom_Quantity.focus();
        return (false);
      }

      var checkOK = "0123456789-";
      var checkStr = theForm.vEcom_Quantity.value;
      var allValid = true;
      var validGroups = true;
      var decPoints = 0;
      var allNum = "";
      for (i = 0; i < checkStr.length; i++) {
        ch = checkStr.charAt(i);
        for (j = 0; j < checkOK.length; j++)
          if (ch == checkOK.charAt(j))
            break;
        if (j == checkOK.length) {
          allValid = false;
          break;
        }
        allNum += ch;
      }
      if (!allValid) {
        alert("Please enter only digit characters in the \"No Seats\" field.");
        theForm.vEcom_Quantity.focus();
        return (false);
      }

      var chkVal = allNum;
      var prsVal = parseInt(allNum);
      if (chkVal != "" && !(prsVal >= 2 && prsVal <= 999)) {
        alert("Please enter a value greater than or equal to \"2\" and less than or equal to \"999\" in the \"No Seats\" field.");
        theForm.vEcom_Quantity.focus();
        return (false);
      }
      return (true);
    }
  </script>
</head>

<body <%=vonloadscript%>>

  <% Server.Execute vShellHi %>

  <table>
    <tr>
      <td style="text-align: center">
<!--        <script>jTitle("<%=fPhraH(000181)%>", 'Basket.jpg')</script>-->
        <img src="../Images/Ecom/MyBasket_<%=svLang %>.png" />
        <h1><!--webbot bot='PurpleText' PREVIEW='My Basket'--><%=fPhra(000181)%></h1>
        <h2><!--webbot bot='PurpleText' PREVIEW='You have chosen to register for the following programs.'--><%=fPhra(000312)%></h2>
      </td>
    </tr>
  </table>

  <table class="table">
    <%
      vCnt = 0
      For i = 1 to svProdMax
        If saProd(2, i) > 0 Then
          vCnt = vCnt + 1
          If vCnt = 1 Then

            '...enter no users
            If vEcom_Media <> "Online" Then
    %>
    <tr>
      <td colspan="5">
        <div style="text-align: center">
          <form method="POST" action="Ecom2Basket.asp" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1">
            <table>
              <tr>
                <td colspan="2" style="text-align: center">
                  <p><%=vTitle%></p>
                </td>
              </tr>

              <% If vEcom_Media = "CDs" Then %>
              <tr>
                <th style="text-align: center">
                  <!--webbot bot='PurpleText' PREVIEW='Quantity'--><%=fPhra(000205)%></th>
                <th style="text-align: center">
                  <!--webbot bot='PurpleText' PREVIEW='Discount'--><%=fPhra(000120)%></th>
              </tr>
              <tr>
                <td style="text-align: center">1 - 9</td>
                <td style="text-align: center">0% </td>
              </tr>
              <tr>
                <td style="text-align: center">10 - 24 </td>
                <td style="text-align: center">10%</td>
              </tr>
              <tr>
                <td style="text-align: center">25 - 49</td>
                <td style="text-align: center">15%</td>
              </tr>
              <tr>
                <td style="text-align: center">50 - 249</td>
                <td style="text-align: center">20%</td>
              </tr>
              <tr>
                <td style="text-align: center">250 +</td>
                <td style="text-align: center">30%</td>
              </tr>
              <% Else %>
              <tr>
                <td style="text-align: center" colspan="2">
                  &nbsp;<!--webbot b-value-required="TRUE" bot="Validation" i-maximum-length="3" i-minimum-length="1" s-data-type="Integer" s-display-name="No Seats" s-number-separators="x" s-validation-constraint="Less than or equal to" s-validation-constraint="Greater than or equal to" s-validation-value="999" s-validation-value="2" -->
                  <input type="text" name="vEcom_Quantity" size="1" value="<%=vEcom_Quantity%>" maxlength="3">
                  <input src="../Images/Buttons/Update_<%=svLang%>.gif" name="I1" type="image">
                </td>
              </tr>
              <% End If %>
            </table>
            <input type="hidden" name="vPage" value="<%=vPage%>">
            <p>&nbsp;</p>
          </form>
        </div>
      </td>
    </tr>
    <% End If %>

    <% If svMembLevel > 3 Then %>
    <tr>
      <td style="text-align: center" colspan="5" class="c3">
        <h3>Additional Discount : <%=Session("Ecom_AdditionalDiscount") & "%" %></h3>
      </td>
    </tr>
    <% End If %>


    <tr>
      <th class="rowshade" rowspan="2" style="width: 15%; text-align: left">
        <!--webbot bot='PurpleText' PREVIEW='Program'--><%=fPhra(000201)%></th>
      <th class="rowshade" rowspan="2" style="width: 50%; text-align: left">
        <!--webbot bot='PurpleText' PREVIEW='Program Title'--><%=fPhra(000320)%></th>
      <th class="rowshade" rowspan="2" style="width: 15%; text-align: center;">
        <!--webbot bot='PurpleText' PREVIEW='Quantity'--><%=fPhra(000205)%></th>
      <th class="rowshade" colspan="2" style="width: 20%; text-align:center;">&nbsp;&nbsp;<!--webbot bot='PurpleText' PREVIEW='Per Seat Cost'--><%=fPhra(000212)%> *</th>
    </tr>
    <tr>
      <th class="rowshade" style="width: 10%; text-align: right">$US&nbsp;&nbsp;&nbsp;</th>
      <th class="rowshade" style="width: 10%; text-align: right">$CA&nbsp;&nbsp;&nbsp;</th>
    </tr>
    <%    
         End If 
    %>
    <form method="POST" action="Ecom2Basket.asp" name="Form<%=i%>">
      <tr>
        <td><%=Mid(saProd(1, i), 10, 7)%></td>
        <td><%=saProd(6, i) & fIf(saProd(0, i) > 0, "<br>[Note: " & saProd(0, i) & "% discount applied]", "") %></td>
        <td style="text-align: center">
          <% If vEcom_Media = "CDs" Then %>
          <select size="1" name="vQty_<%=i%>" onchange="javascript:submitForm(this.form)">
            <% For j = 1 To 24 %>
            <option <%=fselect(j, saprod(2, i))%> value="<%=j%>"><%=j%></option>
            <% Next %> <% For j = 25 To 45 Step 5 %>
            <option <%=fselect(j, saprod(2, i))%> value="<%=j%>"><%=j%></option>
            <% Next %> <% For j = 50 To 250 Step 10 %>
            <option <%=fselect(j, saprod(2, i))%> value="<%=j%>"><%=j%></option>
            <% Next %> <% For j = 275 To 500 Step 25 %>
            <option <%=fselect(j, saprod(2, i))%> value="<%=j%>"><%=j%></option>
            <% Next %> <% For j = 600 To 1000 Step 100 %>
            <option <%=fselect(j, saprod(2, i))%> value="<%=j%>"><%=j%></option>
            <% Next %>
          </select>
          <% Else %> <%=saProd(2, i)%> <% End If %>
        </td>
        <td style="text-align: right"><%=fFormatDecimals(FormatNumber(saProd(3, i) * (1 - saProd(0, i) / 100) ,2))%></td>
        <td style="text-align: right"><%=fFormatDecimals(FormatNumber(saProd(4, i) * (1 - saProd(0, i) / 100) ,2))%></td>
      </tr>
    </form>
    <%
          End If
        Next
  
        '...nothing in the basket?  empty basket in case lines were removed, thus still there with qty = 0
        If vCnt = 0 Then
          svProdNo  = 0
          svProdMax = 0
          ReDim Preserve saProd(6, svProdNo)
          '...save basket values
          Session("ProdNo")    = svProdNo
          Session("ProdMax")   = svProdMax
          Session("Prod")      = saProd       
    %>
    <tr>
      <td colspan="5" style="text-align: center">
        <h6><br><br>
          <!--webbot bot='PurpleText' PREVIEW='No items have been selected, please click <b>Return</b>.'--><%=fPhra(000302)%><br><br>
          <input type="button" onclick="location.href = '<%=vPage%>'" value="<%=bReturn%>" name="bReturn0" class="button">
          <br>to Catalogue</h6>
      </td>
    </tr>
    <% 
        Else
    %>
    <tr>
      <td colspan="5" style="text-align: center;">
        <br />

        <table class="table" style="width: 50%; margin: auto;">
          <tr>
            <td style="text-align: center">
              <input type="button" onclick="location.href = 'Ecom2Programs.asp?vCatlId=<%=Session("Ecom_Catl")%>'" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button"><br>
              <!--webbot bot='PurpleText' PREVIEW='to Catalogue'--><%=fPhra(000333)%>
            </td>
            <td style="text-align: center">
              <input type="button" onclick="location.href = 'Ecom2Basket.asp?vAction=Clear&vPage=<%=vPage%>'" value="<%=bClear%>" name="bClear" class="button"><br>
              <!--webbot bot='PurpleText' PREVIEW='my Basket'--><%=fPhra(000181)%>
            </td>
            <td style="text-align: center">
              <input type="button" onclick="top.main.location.href = 'Ecom2Customer.asp'" value="<%=bContinue%>" name="bContinue" class="button"><br>
              <!--webbot bot='PurpleText' PREVIEW='with Purchase'--><%=fPhra(000334)%>
            </td>
          </tr>
        </table>

        <h3>*
          <!--webbot bot='PurpleText' PREVIEW='Applicable taxes extra for Canadian orders.&nbsp; <br>Products purchases outside Canada are payable in US funds.'--><%=fPhra(000330)%></h3>
      </td>
    </tr>

    <% 
        End If
    %>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


