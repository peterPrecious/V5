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

  If Len(Request("vAdditionalDiscount")) > 0 Then 
    Session("Ecom_AdditionalDiscount") = Request("vAdditionalDiscount")
  ElseIf Len(Session("Ecom_AdditionalDiscount")) = 0 Then
    Session("Ecom_AdditionalDiscount") = 0
  End If 
%>  

<html>

<head>
  <title>Ecom2Programs</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <script>AC_FL_RunContent = 0;</script>
  <script src="/V5/Inc/AC_RunActiveContent.js""></script>
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
<!--              <script>jTitle("<%=fPhraH(000376)%>", 'SingleLicense.jpg')</script>-->
              <% Else %>
              <img src="../Images/Ecom/MultipleLearnerLicense_<%=svLang%>.png" />
<!--              <script>jTitle("<%=fPhraH(000377)%>", 'MultiLicense.jpg')</script>-->
              <% End If %></td>
            <td>
              <table>
                <tr>
                  <th style="text-align: center; padding-bottom: 10px;" colspan="6"><!--webbot bot='PurpleText' PREVIEW='Content Features'--><%=fPhra(001413)%>&nbsp; <span style="font-weight: 400">(<!--webbot bot='PurpleText' PREVIEW='mouseover'--><%=fPhra(001424)%>)</span></th>
                </tr>
                <tr>
                  <td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaAcc.png" title="<!--webbot bot='PurpleText' PREVIEW='Includes compatibility with most screen readers and closed captioning (WCAG Level AA).'--><%=fPhra(001442)%>"></a> </td>
                  <td><!--webbot bot='PurpleText' PREVIEW='Accessible'--><%=fPhra(001415)%></td>
                  <td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaHyb.png" title="<!--webbot bot='PurpleText' PREVIEW='Content available in Flash or HTML.'--><%=fPhra(001628)%>"></a> </td>
                  <td><!--webbot bot='PurpleText' PREVIEW='Hybrid'--><%=fPhra(001613)%></td>
                  <td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaMob.png" title="<!--webbot bot='PurpleText' PREVIEW='Tablet friendly.'--><%=fPhra(001641)%>"></a> </td>
                  <td><!--webbot bot='PurpleText' PREVIEW='Mobile'--><%=fPhra(001416)%></td>
                </tr>
                <tr>
                  <td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaAud.png" title="<!--webbot bot='PurpleText' PREVIEW='Requires headphones or speaker to hear audio.'--><%=fPhra(001443)%>"></a> </td>
                  <td><!--webbot bot='PurpleText' PREVIEW='Audio'--><%=fPhra(001417)%></td>
                  <td>&nbsp;<a href="#"><img border="0" src="../Images/RTE/ModsFeaVid.png" title="<!--webbot bot='PurpleText' PREVIEW='Contains or streams video content.'--><%=fPhra(001445)%>"></a> </td>
                  <td><!--webbot bot='PurpleText' PREVIEW='Video'--><%=fPhra(001418)%></td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
              </table>
            </td>
          </tr>
        </table>


        <h1><!--webbot bot='PurpleText' PREVIEW='Programs'--><%=fPhra(000203)%></h1>
        <p style="text-align:left">
        <!--webbot bot='PurpleText' PREVIEW='To purchase a program, click the <b>Add</b> button and the program will be added to your basket and removed from this list.&nbsp; Click <b>Return</b> to come back to this list to add more programs.'--><%=fPhra(000318)%>&nbsp;<!--webbot bot='PurpleText' PREVIEW='Click on the program title to view the program/module description.'--><%=fPhra(000319)%>
        </p>

        <% If svMembLevel > 3 Then %>
        <form method="POST" action="Ecom2Programs.asp">
          <table style="width:80%; margin:10px auto 0 auto;">
            <tr>
              <td colspan="3"><p>Enter an additional discount that applies to this sale (ie 10%) on top of any other discount that may apply, then click <b>Apply</b><br />.</td>
            </tr>
            <tr>
              <th>Additional Discount: </th>
              <td style="text-align:center; width:70px;"><input type="text" name="vAdditionalDiscount" size="2" value="<%=Session("Ecom_AdditionalDiscount")%>">%</td>
              <td><input type="submit" value="<%=bApply%>" name="bApply" class="button"></td>
            </tr>
          </table>
        </form>
        <% End If %>

        <p class="c6"><% If vCust_EcomDiscOptions > 0 Then %><!--webbot bot='PurpleText' PREVIEW='Note: any applicable discounts will be calculated and displayed in your basket.'--><%=fPhra(000194)%><% End If %> </p>

      </td>
    </tr>
  </table>

  <table class="table">
    <%
      Dim aProgs, aProg, aProg2, vCnt, vBg, vProg_Value, vProg_Name, vProg_US_Lic, vProg_CA_Lic, vConvert
      Dim vFeaturesProg, vFeaturesMods

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
        End If

        
        '...if there is a price override on program table then use that unless price = 9999 or retired
        If Not vProg_Retired And aProg(1) <> 9999 Then
          If vProg_US_Memo > 0 Then aProg(1) = vProg_US_Memo
          If vProg_CA_Memo > 0 Then aProg(2) = vProg_CA_Memo
          If vProg_US_Memo > 0 And vProg_CA_Memo > 0 Then vConvert = False '...don't convert is both values are entered
        End If


        '...modify pricing for group sales unless retired or price is 9999 (inactive) or 1 (free)
        If Not vProg_Retired And aProg(1) <> 1 And aProg(1) <> 9999 And (Session("Ecom_Media") = "Group" Or Session("Ecom_Media") = "Group2") Then

          '...use customer group values unless there is a program group value override
          If vProg_EcomGroupSeat > 0 Then 
            If vProg_EcomGroupSeat = 0.0001 Then 
              vProg_US       = 0.0001
              vProg_CA       = 0.0001
            Else            
              vProg_US       = fRound(aProg(1) * vProg_EcomGroupSeat)
              vProg_CA       = fRound(aProg(2) * vProg_EcomGroupSeat)
            End If
          Else
            vProg_US       = fRound(aProg(1) * vCust_EcomGroupSeat)
            vProg_CA       = fRound(aProg(2) * vCust_EcomGroupSeat)
          End If

          If vProg_EcomGroupLicense > 0 Then 
            '...for set the license to zero, enter 0.0001 in prog table   
            If vProg_EcomGroupLicense = 0.0001 Then 
              vProg_EcomGroupLicense = 0
            End If

            vProg_US_Lic   = fRound(aProg(1) * vProg_EcomGroupLicense)
            vProg_CA_Lic   = fRound(aProg(2) * vProg_EcomGroupLicense)

          Else
            vProg_US_Lic   = fRound(aProg(1) * vCust_EcomGroupLicense)
            vProg_CA_Lic   = fRound(aProg(2) * vCust_EcomGroupLicense)
          End If

        Else

          vProg_US       = aProg(1)
          vProg_CA       = aProg(2)
          vProg_US_Lic   = aProg(1)
          vProg_CA_Lic   = aProg(2)
        End If

        '...convert currency based on vCurrency setting
        If vConvert Then
          If vCust_EcomCurrency = "CA" Then
            If vProg_US     <> 0.0001 Then vProg_US       = fCurrency(vProg_CA)
            If vProg_US_Lic <> 0.0001 Then vProg_US_Lic   = fCurrency(vProg_CA_Lic)
          Else
            If vProg_CA     <> 0.0001 Then vProg_CA       = fCurrency(vProg_US)
            If vProg_CA_Lic <> 0.0001 Then vProg_CA_Lic   = fCurrency(vProg_US_Lic)
          End If
        End If

        If Session("Ecom_Media") = "Online" Then
          vProg_Duration   = aProg(4)
        Else
          vProg_Duration   = 365
        End If

        vProg_Name       = "vProgram" & i


        '...this gets passed through to the basket if item is selected  
        vProg_Value    = Right("00000000" & vCatl_No, 8) & "_" & vProg_Id & "~" & vProg_US & "~" & vProg_CA & "~" & vProg_US_Lic & "~" & vProg_CA_Lic & "~" & vProg_Duration & "~" & fHtmlUnquote(Trim(vProg_Title))

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

        '...display if not retired and price > 0 but not 9999, or not acquired via ecom or on member record   
        sGetProg aProg(0) '...get original record
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
            vBg = "" : If vCnt Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"   '...color every other line       
            
            '...preview hack: make yellow if previewed program
            vBg = fIf(vHiLite=vProg_Id, "bgcolor='yellow' bordercolor='#FFFFFF'", vBg)

            If vCnt = 1 Then '...display title
    %> 
    
    <tr><td colspan="9" style="height:30px;"><h2><%=vCatl_Title%></h2></td></tr>

    <tr>
      <th class="rowshade" style="text-align:left; width:40%;"><!--webbot bot='PurpleText' PREVIEW='Program Title'--><%=fPhra(000320)%></th>
      <th class="rowshade"><!--webbot bot='PurpleText' PREVIEW='Duration Days'--><%=fPhra(000124)%></th>
      <td class="rowshade" style="text-align:center" colspan="2">
        <table class="table">
          <tr><th class="rowshade" colspan="2">&ensp;<!--webbot bot='PurpleText' PREVIEW='Per Seat Cost'--><%=fPhra(000212)%> *&ensp;</th></tr>
          <tr><th class="rowshade">USA$</th><th class="rowshade">CAN$</th></tr>
        </table>
      </td>
      <th class="rowshade" style="text-align:center;"" ><!--webbot bot='PurpleText' PREVIEW='Add to Basket'--><%=fPhra(000321)%></th>
    </tr>

    <%      End If %> 
    
    <%      If vProg_US = 0.0001 Or vProg_US > 1 Then '...Regular line (ie not free) %>

    <form method="POST" action="Ecom2Basket.asp">
      <tr>
        <td style="text-align:left"   <%=vbg%>><p><a <%=fstatx%> href="Ecom2Modules.asp?vProgId=<%=vProg_Id%>"><%=vProg_Title%></a><%=vMods_Features %> <%=fPromo(vProg_Promo)%></p></td>
        <td style="text-align:center" <%=vbg%>><%=vProg_Duration%></td>
        <td style="text-align:center" <%=vbg%>><%=fFormatDecimals(FormatNumber(vProg_US, 2))%></td>
        <td style="text-align:center" <%=vbg%>>&ensp;<%=fFormatDecimals(FormatNumber(vProg_CA, 2))%></td>
	      <td style="text-align:center" <%=vbg%>><input type="submit" value="<%=bAdd%>" name="bAdd" class="button"> <input type="hidden" name="vProgram" value="<%=vProg_Value%>"></td>
      </tr>
      <input type="hidden" name="vPage" value="Ecom2Programs.asp">
    </form>

    <%      Else '...No Charge line    %> 

    <tr>
      <td style="text-align:left"   <%=vbg%>><p><a <%=fstatx%> href="Ecom2Modules.asp?vProgId=<%=vProg_Id%>"><%=vProg_Title%></a><%=vMods_Features %> <%=fPromo(vProg_Promo)%></p></td>
      <td style="text-align:center" <%=vbg%>><%=vProg_Duration%></td>

    <%         If Session("Ecom_Media") = "Online" Then %> 

      <td style="text-align:center" colspan="5" <%=vbg%>>&nbsp;</td>

    <%         Else   %> 

      <td style="text-align:center" colspan="7" <%=vbg%>>&nbsp;</td>
    </tr>

    <% 
              End If 

            End If 

          End If

        End If

      Next

    %>
  </table>
  <br /><br />
  <% If vCnt = 0 Then %>
  <h6><!--webbot bot='PurpleText' PREVIEW='There are no programs available for purchase in the category'--><%=fPhra(000002)%> <b><%=vCatl_Title%></b>.&nbsp;<!--webbot bot='PurpleText' PREVIEW='Previously purchased programs can be accessed by clicking on the &quot;My Content&quot; tab above.&nbsp; If you have just selected them, they can be accessed by clicking on <b>Show My Basket</b> link.'--><%=fPhra(000329)%></h6>
  <% Else %> 
  <h3><!--webbot bot='PurpleText' PREVIEW='Click on the program title to view the program/module description.'--><%=fPhra(000319)%></h3> 
  <h3><a <%=fstatx%> href="Ecom2Basket.asp?vPage=Ecom2Programs.asp" class="c3"><!--webbot bot='PurpleText' PREVIEW='Show My Basket'--><%=fPhra(000239)%></a></h3>
  <h3>* <!--webbot bot='PurpleText' PREVIEW='Applicable taxes extra for Canadian orders.&nbsp; <br>Products purchases outside Canada are payable in US funds.'--><%=fPhra(000330)%></h3>
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

