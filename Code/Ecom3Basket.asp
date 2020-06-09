<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Ecom_Discounts.asp"-->

<%
  Dim aValues, aValue, vTotal, vOnFile, vCnt, vQty, vTitle, vBg, vOk, vTotUS, vTotCA, vTotQty, vStr, vEnd, vOnLoadScript
  Dim vGroup2Rates, aGroup2Rates, aGroup2Rate1, aGroup2Rate2, aGroup2Rate3, aGroup2Rate4, aGroup2Rate5, vMaxSeats

  vEcom_Media = Session("Ecom_Media")
  If fNoValue(vEcom_Media) Then Response.Redirect "EcomError.asp?vMsg=" & Server.UrlEncode("No media selected in " & Request.ServerVariables("Script_Name"))

  '...determine if any discounts apply from the customer file
  sGetCust svCustId

  vMaxSeats = 0
 
  '...these are the values for the various discounts
  vGroup2Rates = fDefault(vCust_EcomGroup2Rates, "5|25~10|45~25|55~50|65~200|75") '...11|10~26|15~100|15~100|15~100|15
  aGroup2Rates = Split(vGroup2Rates, "~")
  aGroup2Rate1 = Split (aGroup2Rates(0), "|") : If (Cint(aGroup2Rate1(0)) > vMaxSeats) Then vMaxSeats = Cint(aGroup2Rate1(0))
  aGroup2Rate2 = Split (aGroup2Rates(1), "|") : If (Cint(aGroup2Rate2(0)) > vMaxSeats) Then vMaxSeats = Cint(aGroup2Rate2(0))
  aGroup2Rate3 = Split (aGroup2Rates(2), "|") : If (Cint(aGroup2Rate3(0)) > vMaxSeats) Then vMaxSeats = Cint(aGroup2Rate3(0))
  aGroup2Rate4 = Split (aGroup2Rates(3), "|") : If (Cint(aGroup2Rate4(0)) > vMaxSeats) Then vMaxSeats = Cint(aGroup2Rate4(0))
  aGroup2Rate5 = Split (aGroup2Rates(4), "|") : If (Cint(aGroup2Rate5(0)) > vMaxSeats) Then vMaxSeats = 500

  '...keep track which form called so can return
  vPage = Request("vPage")

  '...retrieve/create the basket info
  svProdNo       = Session("ProdNo")
  svProdMax      = Session("ProdMax")
  If svProdNo > 0 Then  
    saProd       = Session("Prod")
  Else
    Dim saProd()
    ReDim Preserve saProd (6, svProdNo) '... note: discount % held in 0
  End If

  '...see if clearing basket
  If Request("vAction") = "Clear" Then
    svProdNo  = 0
    svProdMax = 0
    ReDim Preserve saProd (6, svProdNo)
  End If

  '...see if updating a quantity
  For Each vFld In Request.Form
    If Left(vFld, 5) = "vQty_" Then
      i = Cint(Mid(vFld, 6))
      saProd(2, i) = Request(vFld)
      '...add quantity to title
      j = Instr(saProd(6, i), " of ")
      saProd(6, i) = Request(vFld) & Mid(saProd(6, i), j)

      '...check if following lines are freebies and add in the new quantity
      For k = i To Ubound(saProd, 2)
        If saProd(3, k) = 0 Or saProd(4, k) = 0 Then 
          saProd(2, k) = Request(vFld)
         '...add quantity to title
          j = Instr(saProd(6, k), " of ")
          saProd(6, k) = Request(vFld) & Mid(saProd(6, k), j)
        End If
      Next

    End If
  Next

  '...add the selected item into the array
  '...format from Ecom3Programs.asp: <input type="hidden" name="vProgram" value="P1143EN~40~50~40~50~365~Primer on Privacy"></td>

  If Request("vProgram").Count > 0 Then
    vValue  = Replace(Request("vProgram"),"'"," ")
    aValues = Split (vValue,"||")
    For i = 0 To Ubound(aValues)

      aValue  = Split (aValues(i),"~")
      '...remove unnecessary text from program name
      j = Instr(Lcase(aValue(6)), "<br>")
      If j > 0 Then aValue(6) = Left(aValue(6), j-1)
      
      '...ensure this product is not in the array
      vOnFile = False
      If svProdMax > 0 Then
        For j = 1 to Ubound(saProd, 2)
          If saProd(1, j) = aValue(0) Then 
            svProdNo = j
            vOnFile = True
            Exit For
          End If
        Next 
      End If
  
      '...if not in the array, then add item to bottom of array
      If Not vOnFile Then

        '...not sure why we don't set current prodno to max
        svProdMax            = svProdMax + 1

'       svProdNo             = svProdNo + 1
        svProdNo             = svProdMax

        ReDim Preserve saProd (6, svProdNo)
        saProd(0, svProdNo) = 0                                          '...percentage discount
        saProd(1, svProdNo) = aValue(0) 							                   '...catl_no + program Id
        saProd(2, svProdNo) = 1                                          '...start with 1
        saProd(3, svProdNo) = aValue(1)                                  '...US price each
        saProd(4, svProdNo) = aValue(2)                                  '...CA price each
        saProd(5, svProdNo) = 365                                        '...license is always a duration of 365 days
        saProd(6, svProdNo) = "1 of " & aValue(6)                        '...description

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


  '...see if there is a discount
  sCheckDiscounts

  '...save basket values
  Session("ProdNo")  = svProdNo
  Session("ProdMax") = svProdMax
  Session("Prod")    = saProd       
 
%>


<html>

<head>
  <title>Ecom3Basket</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script> 
    function submitForm(theForm)
    {
      theForm.action = 'Ecom3Basket.asp';
      theForm.submit();
    }
  </script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>AC_FL_RunContent = 0;</script>
  <script src="/V5/Inc/AC_RunActiveContent.js" language="javascript"></script>
  <script>
    function jTitle (vTitle, vImage) {
      var vParm = "title=" + vTitle + '&image=/V5/Images/Titles/' + vImage;
      AC_FL_RunContent('codebase','//download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0','name','flashVars','width','265','height','85','align','middle','id','flashVars','src','/V5/Images/Titles/VuTitles','FlashVars',vParm,'quality','high','bgcolor','#ffffff','allowscriptaccess','sameDomain','allowfullscreen','false','pluginspage','///go/getflashplayer','movie','/V5/Images/Titles/VuTitles');
    }
  </script>
  <style>
    .title { text-align: center; background-color: #DDEEF9; border-color: #FFFFFF; white-space: nowrap; }
  </style>
</head>

<body>

  <% Server.Execute vShellHi %>
  <table class="table">
    <tr>
      <td style="text-align: center">
        <div style="text-align: left">
<!--          <script>jTitle("<%=fPhraH(000181)%>", 'Basket.jpg')</script>-->
          <img src="../Images/Ecom/MyBasket_<%=svLang %>.png" />
        </div>
        <h1><!--webbot bot='PurpleText' PREVIEW='My Basket'--><%=fPhra(000181)%></h1>
        <p class="c2"><!--webbot bot='PurpleText' PREVIEW='You have chosen to register for the following programs.&nbsp; Multi-learner discounts apply and will be calculated/shown after the number of seats is entered.'--><%=fPhra(000378)%></p>
        <a href="javascript:toggle('div_discounts');" class="c3"><!--webbot bot='PurpleText' PREVIEW='Show Multi-learner Discounts'--><%=fPhra(000922)%></a>
        <div id="div_discounts" class="div">
          <table style="width: 60%; margin: auto;">
            <tr>
              <td style="text-align: center" height="30" bgcolor="#DDEEF9">
                <p><!--webbot bot='PurpleText' PREVIEW='The following multi-learner discounts apply.'--><%=fPhra(000381)%>
                </p>
              </td>
            </tr>
            <tr>
              <td style="text-align: center">
                <table border="0" id="table8" bordercolor="#DDEEF9" cellpadding="0">
                  <tr>
                    <th nowrap style="text-align: center" colspan="3">
                      <!--webbot bot='PurpleText' PREVIEW='Total Seats'--><%=fPhra(000282)%>&nbsp;&nbsp; </th>
                    <th nowrap style="text-align: center">&nbsp;<!--webbot bot='PurpleText' PREVIEW='Discount'--><%=fPhra(000120)%></th>
                  </tr>
                  <tr>
                    <td style="text-align:right"><%=aGroup2Rate1(0)%></td>
                    <td style="text-align: center">-</td>
                    <td><%=aGroup2Rate2(0) - 1%></td>
                    <td style="text-align: center"><%=aGroup2Rate1(1) & "%" %></td>
                  </tr>
                  <% If Cint(aGroup2Rate2(1)) > Cint(aGroup2Rate1(1)) Then %>
                  <tr>
                    <td style="text-align:right"><%=aGroup2Rate2(0)%></td>
                    <td style="text-align: center">-</td>
                    <td><%=aGroup2Rate3(0) - 1%></td>
                    <td style="text-align: center"><%=aGroup2Rate2(1) & "%" %></td>
                  </tr>
                  <% End If %>
                  <% If Cint(aGroup2Rate3(1)) > Cint(aGroup2Rate2(1)) Then %>
                  <tr>
                    <td style="text-align:right"><%=aGroup2Rate3(0)%></td>
                    <td style="text-align: center">-</td>
                    <td><%=aGroup2Rate4(0) - 1%></td>
                    <td style="text-align: center"><%=aGroup2Rate3(1) & "%" %></td>
                  </tr>
                  <% End If %>
                  <% If Cint(aGroup2Rate4(1)) > Cint(aGroup2Rate3(1)) Then %>
                  <tr>
                    <td style="text-align:right"><%=aGroup2Rate4(0)%></td>
                    <td style="text-align: center">-</td>
                    <td><%=aGroup2Rate5(0) - 1%></td>
                    <td style="text-align: center"><%=aGroup2Rate4(1) & "%" %></td>
                  </tr>
                  <% End If %>
                  <% If Cint(aGroup2Rate5(1)) > Cint(aGroup2Rate4(1)) Then %>
                  <tr>
                    <td style="text-align:right"><%=aGroup2Rate5(0)%></td>
                    <td style="text-align: center">-</td>
                    <td>500</td>
                    <td style="text-align: center"><%=aGroup2Rate5(1) & "%" %></td>
                  </tr>
                  <% End If %>
                </table>
                <br />
                <div class="red"><a class="red" href="mailto:<%=fDefault(vCust_Email, "info@vubiz.com")%>?subject=Multi User License Enquiry"><b><!--webbot bot='PurpleText' PREVIEW='Email us'--><%=fPhra(001220)%></a> <!--webbot bot='PurpleText' PREVIEW='if you are interested in more seats.'--><%=fPhra(001603)%></div>
 
              </td>
            </tr>
          </table>
          <br>
        </div>
        <h3 style="margin-bottom: 10px;"><!--webbot bot='PurpleText' PREVIEW='Please enter the number of seats you wish to assign to each Program.'--><%=fPhra(000471)%></h3>
      </td>
    </tr>
  </table>
  <table class="table">
    <%
      vCnt = 0 : vQty = 0
      For i = 1 to svProdMax
        If saProd(2, i) > 0 Then
          vCnt = vCnt + 1
          If saProd(3, i) > 0 Then vQty = vQty + saProd(2, i)
          If vCnt = 1 Then
    %>
    <tr>
      <th class="rowshade" rowspan="2" style="width: 15%; text-align: left"><!--webbot bot='PurpleText' PREVIEW='Program'--><%=fPhra(000201)%></th>
      <th class="rowshade" rowspan="2" style="width: 50%; text-align: left"><!--webbot bot='PurpleText' PREVIEW='Program Title'--><%=fPhra(000320)%></th>
      <th class="rowshade" rowspan="2" style="width: 15%; text-align:center;"><!--webbot bot='PurpleText' PREVIEW='No. of <br>Seats'--><%=fPhra(000324)%></th>
      <th class="rowshade" colspan="2" style="width: 20%; text-align:center;">&nbsp;&nbsp;<!--webbot bot='PurpleText' PREVIEW='Per Seat Cost'--><%=fPhra(000212)%> *</th>
    </tr>
    <tr>
      <th class="rowshade" style="width: 10%; text-align: right">$US&nbsp;&nbsp;&nbsp;</th>
      <th class="rowshade" style="width: 10%; text-align: right">$CA&nbsp;&nbsp;&nbsp;</th>
    </tr>
    <%    
         End If
  
         vBg = "" : If vCnt Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"   '...color every other line       
    %>
    <form method="POST" action="Ecom3Basket.asp" name="Form<%=i%>">
      <tr>
        <td><%=Mid(saProd(1, i), 10, 7)%></td>
        <td><%=saProd(6, i)%></td>
        <td style="text-align: center; white-space: nowrap">
          <!-- if this is a freebie line then display quantities -->
          <% If saProd(3, i) = 0 Or saProd(4, i) = 0 Then %>
          <%=saProd(2, i) %>
          <% Else %>
          <%   If svMembLevel < 6 Then %>
          <select size="1" name="vQty_<%=i%>" onchange="javascript:submitForm(this.form)">
            <% For j = 1 To vMaxSeats %>
            <option <%=fselect(j, saprod(2, i))%> value="<%=j%>"><%=j%></option>
            <% Next %>
          </select>
          <%   Else %>
          <input type="text" name="vQty_<%=i%>" onblur="javascript:submitForm(this.form)" size="4" value="<%=saprod(2, i)%>">
          <%   End If %>
          <% End If %>
        </td>
        <td style="text-align: right; white-space: nowrap"><%=FormatNumber(saProd(3, i) * (1 - saProd(0, i) / 100), 2)%></td>
        <td style="text-align: right; white-space: nowrap"><%=FormatNumber(saProd(4, i) * (1 - saProd(0, i) / 100), 2)%></td>
      </tr>
      <input type="hidden" name="vPage" value="Ecom3Programs.asp">
    </form>
    <%
        End If
      Next
  
        '...nothing in the basket?  empty basket in case lines were removed, thus still there with qty = 0
        If vCnt = 0 Then
          svProdNo  = 0
          svProdMax = 0
          ReDim Preserve saProd (6, svProdNo)
          '...save basket values
          Session("ProdNo")    = svProdNo
          Session("ProdMax")   = svProdMax
          Session("Prod")      = saProd       
    %>
    <tr>
      <td colspan="5" style="text-align: center;">
        <h6><br><br>
          <!--webbot bot='PurpleText' PREVIEW='No items have been selected, please click <b>Return</b>.'--><%=fPhra(000302)%><br><br>
          <input type="button" onclick="location.href = '<%=vPage%>'" value="<%=bReturn%>" name="bReturn0" class="button">
          <br>
          <!--webbot bot='PurpleText' PREVIEW='to Catalogue'--><%=fPhra(000333)%></h6>
      </td>
    </tr>
    <% 
        Else
    %>
    <tr>
      <td colspan="5" style="text-align: center;"><br />

        <table class="table" style="width:50%; margin:auto;">
          <tr>
            <td style="text-align: center">
              <input type="button" onclick="location.href = '<%=vPage%>'" value="<%=bReturn%>" name="bReturn" id="bReturn" class="button"><br><!--webbot bot='PurpleText' PREVIEW='to Catalogue'--><%=fPhra(000333)%>
            </td>
            <td style="text-align: center">
              <input type="button" onclick="location.href = 'Ecom3Basket.asp?vAction=Clear&vPage=<%=vPage%>'" value="<%=bClear%>" name="bClear" class="button"><br><!--webbot bot='PurpleText' PREVIEW='my Basket'--><%=fPhra(000181)%>
            </td>
            <td style="text-align: center">
              <input type="button" onclick="top.main.location.href = 'Ecom2Customer.asp'" value="<%=bContinue%>" name="bContinue" class="button"><br><!--webbot bot='PurpleText' PREVIEW='with Purchase'--><%=fPhra(000334)%>
            </td>
          </tr>
        </table>

        <h3>* 
          <!--webbot bot='PurpleText' PREVIEW='Applicable taxes extra for Canadian orders.&nbsp; <br>Products purchases outside Canada are payable in US funds.'--><%=fPhra(000330)%>&nbsp;&nbsp;  
          <!--webbot bot='PurpleText' PREVIEW='This transaction will be processed through a Canadian financial institution. Please refer to your credit card agreement regarding fees, if any.'--><%=fPhra(001837)%>
        </h3>
      </td>
    </tr>
    <% 
        End If
    %>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


