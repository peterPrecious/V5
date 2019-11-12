<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Querystring.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->
<!--#include virtual = "V5/Inc/Db_Mods.asp"-->

<%

  '...this hack grabs the vMemo field for Previews where we want to highlight the program previewed/promoted
  '   it is only used if accompanied by vTraining
  Dim vHiLite  
  sGetQueryString 
  vHiLite = fIf (Len(vTraining) > 0, vMemo, "")

  Dim vGroup2Rates, aGroup2Rates, aGroup2Rate1, aGroup2Rate2, aGroup2Rate3, aGroup2Rate4, aGroup2Rate5
  
  If Len(Request("vAdditionalDiscount")) > 0 Then 
    Session("Ecom_AdditionalDiscount") = Request("vAdditionalDiscount")
  ElseIf Len(Session("Ecom_AdditionalDiscount")) = 0 Then
    Session("Ecom_AdditionalDiscount") = 0
  End If  

  '...determine if any discounts apply from the customer file
  sGetCust svCustId

  vGroup2Rates = fDefault(vCust_EcomGroup2Rates, "5|25~10|45~25|55~50|65~200|75")
  aGroup2Rates = Split(vGroup2Rates, "~")
  aGroup2Rate1 = Split (aGroup2Rates(0), "|")
  aGroup2Rate2 = Split (aGroup2Rates(1), "|")
  aGroup2Rate3 = Split (aGroup2Rates(2), "|")
  aGroup2Rate4 = Split (aGroup2Rates(3), "|")
  aGroup2Rate5 = Split (aGroup2Rates(4), "|")

%>

<html>

<head>
  <title>Ecom3Programs</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>AC_FL_RunContent = 0;</script>
  <script src="/V5/Inc/AC_RunActiveContent.js"></script>
  <script src="/v5/Inc/Functions.js"></script>
  <script>
    function jTitle (vTitle, vImage) {
      var vParm = "title=" + vTitle + '&image=/V5/Images/Titles/' + vImage;
      AC_FL_RunContent('codebase','//download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0','name','flashVars','width','265','height','85','align','middle','id','flashVars','src','/V5/Images/Titles/VuTitles','FlashVars',vParm,'quality','high','bgcolor','#ffffff','allowscriptaccess','sameDomain','allowfullscreen','false','pluginspage','///go/getflashplayer','movie','/V5/Images/Titles/VuTitles');
    }
  </script>
  <style>
    .title { text-align: center; background-color: #DDEEF9; border-color: #FFFFFF; white-space: nowrap; }
    th, td { padding: 2px; }
  </style>
</head>

<body>

  <% Server.Execute vShellHi %>

  <table class="table">
    <tr>
      <td>
        <table>
          <tr>
            <td><% If Session("Ecom_Media") = "Online" Then  %>
              <img src="../Images/Ecom/SingleLearnerLicense_<%=svLang%>.png" />
<!--              <script>jTitle("/*--{[--*/Single Learner License/*--]}--*/", 'SingleLicense.jpg')</script>-->
              <% Else %>
              <img src="../Images/Ecom/MultipleLearnerLicense_<%=svLang%>.png" />
<!--              <script>jTitle("/*--{[--*/Multiple Learner License/*--]}--*/", 'MultiLicense.jpg')</script>-->
              <% End If %></td>
            <td>
              <table>
                <tr>
                  <th style="text-align: center; padding-bottom: 10px;" colspan="6"><!--[[-->Content Features<!--]]-->&nbsp; <span style="font-weight: 400">(<!--[[-->mouseover<!--]]-->)</span></th>
                </tr>
                <tr>
                  <td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaAcc.png" title="<!--[[-->Includes compatibility with most screen readers and closed captioning (WCAG Level AA).<!--]]-->"></a> </td>
                  <td><!--[[-->Accessible<!--]]--></td>
                  <td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaHyb.png" title="<!--[[-->Content available in Flash or HTML.<!--]]-->"></a> </td>
                  <td><!--[[-->Hybrid<!--]]--></td>
                  <td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaMob.png" title="<!--[[-->Tablet friendly.<!--]]-->"></a> </td>
                  <td><!--[[-->Mobile<!--]]--></td>
                </tr>
                <tr>
                  <td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaAud.png" title="<!--[[-->Requires headphones or speaker to hear audio.<!--]]-->"></a> </td>
                  <td><!--[[-->Audio<!--]]--></td>
                  <td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaVid.png" title="<!--[[-->Contains or streams video content.<!--]]-->"></a> </td>
                  <td><!--[[-->Video<!--]]--></td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
              </table>
            </td>
          </tr>
        </table>

        <h1>
          <!--[[-->Programs<!--]]--></h1>

        <p class="c3" style="text-align: left; padding-bottom: 10px;">
          <!--[[-->To purchase a program, click the <b>Add</b> button and the program will be added to your basket and removed from this list.&nbsp; Click <b>Return</b> to come back to this list to add more programs.<!--]]-->
          <!--[[-->Click on the program title to view the program/module description.<!--]]-->&nbsp;&nbsp;
          <!--[[-->Note: The prices below are individual seat prices.&nbsp; Multi-learner discounts are applied in the next step when the number of seats is entered.<!--]]-->
          <!--[[--><b>Terms of License</b>: Your license is active for one year (365 days) from the date of purchase.<!--]]-->
          <!--[[-->As the Facilitator of this Web based service, you can take the purchased course(s) yourself at no additional charge.<!--]]-->
        </p>

      </td>
    </tr>
    <tr>
      <td style="text-align: center;"><a href="javascript:toggle('div_discounts');" class="c3"><!--[[-->Show Multi-learner Discounts<!--]]--></a>&ensp;&ensp;
      <div id="div_discounts" class="div">
        <table class="table">
          <tr>
            <td style="text-align: center; height: 30px; background-color: #DDEEF9"><!--[[-->The following multi-learner discounts apply.<!--]]--></td>
          </tr>
          <tr>
            <td style="text-align: center">
              <table style="width: 400px; margin: auto;">
                <tr>
                  <th style="text-align: center" colspan="3"><!--[[-->Total Seats<!--]]--></th>
                  <th style="text-align: center"><!--[[-->Discount<!--]]--></th>
                </tr>
                <tr>
                  <td style="text-align: right"><%=aGroup2Rate1(0)%></td>
                  <td style="text-align: center">-</td>
                  <td><%=aGroup2Rate2(0) - 1%></td>
                  <td style="text-align: center"><%=aGroup2Rate1(1) & "%" %></td>
                </tr>
                <% If Cint(aGroup2Rate2(1)) > Cint(aGroup2Rate1(1)) Then %>
                <tr>
                  <td style="text-align: right"><%=aGroup2Rate2(0)%></td>
                  <td style="text-align: center">-</td>
                  <td><%=aGroup2Rate3(0) - 1%></td>
                  <td style="text-align: center"><%=aGroup2Rate2(1) & "%" %></td>
                </tr>
                <% End If %>
                <% If Cint(aGroup2Rate3(1)) > Cint(aGroup2Rate2(1)) Then %>
                <tr>
                  <td style="text-align: right"><%=aGroup2Rate3(0)%></td>
                  <td style="text-align: center">-</td>
                  <td><%=aGroup2Rate4(0) - 1%></td>
                  <td style="text-align: center"><%=aGroup2Rate3(1) & "%" %></td>
                </tr>
                <% End If %>
                <% If Cint(aGroup2Rate4(1)) > Cint(aGroup2Rate3(1)) Then %>
                <tr>
                  <td style="text-align: right"><%=aGroup2Rate4(0)%></td>
                  <td style="text-align: center">-</td>
                  <td><%=aGroup2Rate5(0) - 1%></td>
                  <td style="text-align: center"><%=aGroup2Rate4(1) & "%" %></td>
                </tr>
                <% End If %>
                <% If Cint(aGroup2Rate5(1)) > Cint(aGroup2Rate4(1)) Then %>
                <tr>
                  <td style="text-align: right"><%=aGroup2Rate5(0)%></td>
                  <td style="text-align: center">-</td>
                  <td>500</td>
                  <td style="text-align: center"><%=aGroup2Rate5(1) & "%" %></td>
                </tr>
                <% End If %>
              </table>
              <% If svMembLevel > 3 Then %>
              <form method="POST" action="Ecom3Programs.asp">
                <table style="width: 400px; margin: auto;">
                  <tr>
                    <td style="text-align: center" colspan="3">
                      <p>
                        Enter an additional discount that applies to this sale
                        <br>
                        (ie 10%) on top of the discount above, then click <b>Apply</b>.
                      </p>
                    </td>
                  </tr>
                  <tr>
                    <td style="text-align: right">Additional Discount: </td>
                    <td style="text-align: center">
                      <input type="text" name="vAdditionalDiscount" size="2" value="<%=Session("Ecom_AdditionalDiscount")%>">%</td>
                    <td>
                      <input type="submit" value="<%=bApply%>" name="bApply" class="button"></td>
                  </tr>
                </table>
              </form>
              <% End If %>
              <p>
                <a href="mailto:<%=fDefault(vCust_Email, "info@vubiz.com")%>?subject=Multi User License Enquiry"><b><font color="#FF0000"><!--[[-->Email us<!--]]--></font></b></a><font color="#FF0000">&ensp;<!--[[-->if you are interested in more seats.<!--]]--></font>
            </td>
          </tr>
        </table>
        <br>
      </div>
        <a <%=fstatx%> href="Ecom3Basket.asp?vPage=Ecom3Programs.asp" class="c3"><!--[[-->Show My Basket<!--]]--></a><br>&nbsp;
      </td>
    </tr>
  </table>

  <table class="table">

    <%
      Dim aProgs, aProg, aProg2, vCnt, vBg, vProg_Value, vProg_Name, vProg_US_Lic, vProg_CA_Lic, vConvert

      vConvert = True '...default is to enter one price and convert the other unless both prices are entered in the program table, then do not convert

      If Request("vCatlNo").Count > 0 Then
        Session("Ecom_Catl")  = Request("vCatlNo")
      End If
    
      Session("Ecom_Prog")  = ""
      Session("Ecom_Mods")  = ""
      
      Dim aCatl, vCatl, vOk
      
      '...get customer product string
      sGetCust svCustId

      '...get any user and ecom program 
      If svSecure Then 
        sGetMemb svMembno
        vEcom_Programs = fEcomPrograms (svCustId, svMembId)
      End If

      '...if no catalogue id passed, use the first catl id
      If Len(Session("Ecom_Catl")) = 0 Then 
        sGetCatl_Rs svCustId
        If Not oRs2.Eof Then
          sReadCatl
          Session("Ecom_Catl") = vCatl_No
        End If
      End If
      
      sGetCatl Session("Ecom_Catl")
    
      '...retrieve basket info to refine descriptions if In Basket
      svProdNo       = Session("ProdNo")
      svProdMax      = Session("ProdMax")
      If svProdNo > 0 Then  
        Dim saProd
        saProd       = Session("Prod")
      End If

      vCnt = 0

      '...get the program strings from the customer content string
      aProgs = Split(vCatl_Programs)

      For i = 0 To Ubound(aProgs) '...aProgs(i): "P1001EN~50~79~23.5~90"

        vConvert = True '...default is to enter one price and convert the other unless both prices are entered in the program table, then do not convert
        aProg = Split(aProgs(i), "~") 
        sGetProg aProg(0)

        '...get the feature set of the first module - ensure it's a program that contains modules as some are "placeholders" 
        sGetMods Left(vProg_Mods, 6)  

        '...get mod features but if this program is a placeholder grab the one next on the list
        If Len(Trim(vProg_Mods)) = 0 And uBound(aProgs) > i Then 
          aProg = Split(aProgs(i+1), "~") 
          sGetProg aProg(0)
          sGetMods Left(vProg_Mods, 6)  
          '...get the original program
          aProg = Split(aProgs(i), "~") 
          sGetProg aProg(0)
          sGetMods Left(vProg_Mods, 6)  
        End If
  
        
        '...if there is a price override on program table then use that unless price = 9999 or retired
        If Not vProg_Retired And aProg(1) <> 9999 Then
          If vProg_US_Memo > 0 Then aProg(1) = vProg_US_Memo
          If vProg_CA_Memo > 0 Then aProg(2) = vProg_CA_Memo
          If vProg_US_Memo > 0 And vProg_CA_Memo > 0 Then vConvert = False '...don't convert is both values are entered
        End If

        vProg_US       = aProg(1)
        vProg_CA       = aProg(2)

        '...convert currency based on vCurrency setting (unless vConvert = False)
        If vConvert Then
          If vCust_EcomCurrency = "CA" Then
            If vProg_US     <> 0.0001 Then vProg_US       = fCurrency(vProg_CA)
          Else
            If vProg_CA     <> 0.0001 Then vProg_CA       = fCurrency(vProg_US)
          End If
        End If

        If Session("Ecom_Media") = "Online" Then
          vProg_Duration   = aProg(4)
        Else
          vProg_Duration   = 365
        End If

        vProg_Name       = "vProgram" & i

        '...this gets passed through to the basket if item is selected  
        vProg_Value    = Right("00000000" & vCatl_No, 8) & "_" & vProg_Id & "~" & vProg_US & "~" & vProg_CA & "~0~0~" & vProg_Duration & "~" & fHtmlUnquote(Trim(vProg_Title))

        '...see if next product is included no charge, if so, add to the vProg_Value
        If vProg_US > 1 Or vProg_US = 0.0001 Then
          For j = i + 1 To Ubound(aProgs)
            aProg2 = Split(aProgs(j), "~") 
            sGetProg aProg2(0)
            If aProg2(1) = 1 Then
              vProg_Value = vProg_Value & "||" & Right("00000000" & vCatl_No, 8) & "_" & vProg_Id & "~0~0~0~0~" & vProg_Duration & "~" & fHtmlUnquote(Trim(vProg_Title))
            Else
              Exit For
            End If
          Next
        End If

        '...display if price > 0 but not 9999 or retired, or not acquired via ecom or on member record   
        sGetProg aProg(0) '...get original record
'       If vProg_US > 0 And vProg_US <> 9999 And vProg_CA <> 9999 And Instr(vEcom_Programs, vProg_Id) = 0 And Instr(vMemb_Programs, vProg_Id) = 0 Then
        If Not vProg_Retired And vProg_US > 0 And vProg_US <> 9999 And vProg_CA <> 9999 And Instr(vEcom_Programs, vProg_Id) = 0 And Instr(vMemb_Programs, vProg_Id) = 0 Then
       
          '...only display if NOT already in basket
          vOk = True
          For j = 1 to svProdMax
           If saProd(1, j) = vProg_Id Then 
             vOk = False
             Exit For
           End If
          Next

          If vOk Then  '...ie, if NOT in basket then display
            vCnt = vCnt + 1
            vBg = "" : If vCnt Mod 2 = 0 Then vBg = "background-color:#DDEEF9; border-color:#FFFFFF;"   '...color every other line       
            '...preview hack: make yellow if previewed program
            vBg = fIf(vHiLite=vProg_Id, "background-color:yellow; border-color:#FFFFFF", vBg)

            If vCnt = 1 Then '...display title
    %>

    <tr>
      <td colspan="4" style="height: 30px;"><h2><%=vCatl_Title%></h2></td>
    </tr>

    <tr>
      <th class="rowshade" style="width: 50%; text-align: left;"><!--[[-->Program Title<!--]]--></th>
      <th class="rowshade" colspan="2" style="width: 20%;">
        <table class="table">
          <tr>
            <th class="rowshade" style="text-align: right;" colspan="2"><!--[[-->Per Seat Cost<!--]]--><b>*</b></th>
          </tr>
          <tr>
            <th class="rowshade">$USA&nbsp; </th>
            <th class="rowshade">&nbsp;$CAN&nbsp; </th>
          </tr>
        </table>
      </th>
      <th class="rowshade" style="width: 20%; text-align: center;"><!--[[-->Add to Basket<!--]]--></th>
    </tr>

    <%      
            End If
    
            If vProg_US = 0.0001 Or vProg_US > 1 Then '...Regular line (ie not free) 
    %>

    <tr>
      <td style="<%=vBg%>"><a <%=fstatx%> href="Ecom2Modules.asp?vProgId=<%=vProg_Id%>"><%=vProg_Title%></a><%=vMods_Features %> <%=fPromo(vProg_Promo)%></td>
      <td style="text-align: right; <%=vBg%>"><%=FormatNumber(vProg_US, 2) %></td>
      <td style="text-align: right; <%=vBg%>">&ensp;<%=FormatNumber(vProg_CA, 2) %></td>
      <td style="text-align: center; <%=vBg%>">
        <form method="POST" action="Ecom3Basket.asp">
          <input type="submit" value="<%=bAdd%>" name="bAdd" class="button">
          <input type="hidden" name="vProgram" value="<%=vProg_Value%>">
          <input type="hidden" name="vPage" value="Ecom3Programs.asp">
        </form>
      </td>
    </tr>

    <%       Else '...No Charge line %>

    <tr>
      <td style="<%=vBg%>">
        <p><a <%=fstatx%> href="Ecom2Modules.asp?vProgId=<%=vProg_Id%>"><%=vProg_Title%></a><%=vMods_Features %> <%=fPromo(vProg_Promo)%></p>
      </td>
      <td style="text-align: center; <%=vBg%>" colspan="2">&nbsp;</td>
      <td style="text-align: center; <%=vBg%>">&nbsp;</td>
    </tr>

    <% 
            End If 

          End If

        End If

      Next

    %>
  </table>


  <% If vCnt = 0 Then %>
  <h6>
    <!--[[-->There are no programs available for purchase in the category<!--]]-->
    <b><%=vCatl_Title%></b>.&nbsp;
    <!--[[-->Previously purchased programs can be accessed by clicking on the &quot;My Content&quot; tab above.&nbsp; If you have just selected them, they can be accessed by clicking on <b>Show My Basket</b> link.<!--]]-->
  </h6>
  <% Else %>
  <h3>
    <!--[[-->Click on the program title to view the program/module description.<!--]]--></h3>
  <h3><b>* </b>
    <!--[[-->Applicable taxes extra for Canadian orders.&nbsp;
    <br>
    Products purchases outside Canada are payable in US funds.<!--]]-->
  </h3>
  <% End If %>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


<%
  '...this rounds up values to end in one of: 10, 15 or 19
  '   if there's something to round up and currency is other than equal.
  Function fRound (i)
    If i <> 0 And vCurrency <> 1 Then
      If i > 10 Then
        i = (Round(i) / 10) * 10
        '...round to 5 if 1-4 or 9 if 6-8
        If Right(i, 1) > 0 And Right(i, 1) < 6 Then
          i = Left(i, len(i)-1) & "5"
        ElseIf Right(i, 1) > 5 And Right(i, 1) < 9 Then 
          i = Left(i, len(i)-1) & "9"
        End If
      Else
        i = 10  
      End If 
    End If
    fRound = i
  End Function

  '...This converts currency
  Function fCurrency (i)
    If IsNumeric(i) Then
      If i = 1 Then
        fCurrency = 1
      ElseIf i = 0 Or i = 0.0001 Then
        fCurrency = 0
      ElseIf i > 0 Then
        If vCust_EcomCurrency = "CA" Then
          fCurrency = i * vCurrency
        Else
          fCurrency = i / vCurrency
        End If
        fCurrency = fRound(fCurrency)
      End If
    Else
      fCurrency = 0
    End If
  End Function
%>