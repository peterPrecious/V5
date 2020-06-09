<!--#include virtual = "V5/Inc/Setup.asp"-->
<% vBypassSecurity = True %>
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Db_Ecom.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<% 
  Session("Ecom_Catl")  = ""
  Session("Ecom_Prog")  = ""
  Session("Ecom_Mods")  = ""
%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <title>Vubiz Catalogue</title>
  <base target="_self">
  <script>AC_FL_RunContent = 0;</script>
  <script src="/V5/Inc/AC_RunActiveContent.js"></script>
  <script>
      function jTitle (vTitle, vImage) {
        var vParm = "title=" + vTitle + '&image=/V5/Images/Titles/' + vImage;
        AC_FL_RunContent('codebase','//download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=6,0,0,0','name','flashVars','width','265','height','85','align','middle','id','flashVars','src','/V5/Images/Titles/VuTitles','FlashVars',vParm,'quality','high','bgcolor','#ffffff','allowscriptaccess','sameDomain','allowfullscreen','false','pluginspage','///go/getflashplayer','movie','/V5/Images/Titles/VuTitles');
      }
  </script>
</head>

<body>

  <% Server.Execute vShellHi %>

  <table class="table">
    <tr>
      <td style="text-align:center">
        <img src="../Images/Ecom/Categories_<%=svLang %>.png" />
<!--        <script>jTitle("<%=fPhraH(000086)%>", 'Categories.jpg')</script>-->
        <h1><!--webbot bot='PurpleText' PREVIEW='Categories'--><%=fPhra(000086)%></h1>
        <h3 class="c3"><!--webbot bot='PurpleText' PREVIEW='Click on any category title below and a list of learning programs available for purchase will appear on the right.'--><%=fPhra(000099)%>&nbsp;<!--webbot bot='PurpleText' PREVIEW='Sort by'--><%=fPhra(000243)%>...</h3>
        <p style="text-align:center">
          <a href="Ecom2Catalogue.asp?vSort=n" class="c3"><!--webbot bot='PurpleText' PREVIEW='Default Order'--><%=fPhra(000310)%></a>&nbsp; 
          <a href="Ecom2Catalogue.asp?vSort=y" class="c3"><!--webbot bot='PurpleText' PREVIEW='Category Order'--><%=fPhra(000311)%></a><br />
          <br />
        </p>
      </td>
    </tr>
  </table>

  <table class="table">
    <%
      Dim aCatl, vCnt, vBg, vOk, aProgs, aProg, vSort, vCustId, vInitCatlNo

      sGetCust svCustId 

      '...if addon then grab the parent's catalogue else grab this account's catalogue
      If Session("Ecom_Media") = "AddOn2" Then
        vCustId = Left(vCust_Id, 4) & vCust_ParentId
      Else
        vCustId = svCustId
      End If

      '...if secure then get the member and ecom info
      If svSecure Then
        sGetMemb svMembno
        vEcom_Programs = fEcomPrograms (svCustId, svMembId)
      End If

      vSort = fDefault(Request("vSort"), "n")

      '...get the Catalogue info
      If vSort = "y" Then 
        sGetCatlByTitle_Rs vCustId
      Else
        sGetCatl_Rs vCustId
      End If

      vCnt = 0
      Do While Not oRs2.Eof

        sReadCatl
        vOk = True

        aProgs = Split(vCatl_Programs)   '...aProgs(j): "P1001EN~50~79~23.5~90"

        For j = 0 To Ubound(aProgs) 

          '...check each catalogue item to ensure it contains programs that are available for purchase
          aProg = Split(aProgs(j), "~") 

          '...get pricing unless price is 9999  
          vProg_Id       = aProg(0)
          vProg_US       = aProg(1)
          vProg_CA       = aProg(2)
          vProg_Duration = aProg(4)

          '...must not be inactive, free program have a duration of zero (forever free)
          If vProg_US <> 9999 And vProg_US <> 0 And vProg_Duration > 0 Then          

            '...if signed in, ensure it hasn't been purchased or has expired on then member file
            If Not svSecure Then
              vOk = True
              Exit For   '...as long as there's one program available, then display the group            


            Else
   
              '...if on user table they cannot be purchased
              If Instr(vMemb_Programs, vProg_Id) > 0 Then
              
                '...any expirey date on user record?
                If fFormatDate(vMemb_Expires) <> " " Then
                  vMemb_Expires = vMemb_Expires
                '...else a duration  
                Else
                  vMemb_Expires = DateAdd("d", vMemb_Duration, svMembFirstVisit)
                End If

                If DateDiff("d", Now, vMemb_Expires) < 1 Then '...ensure member program has expired so he can buy it
                  vOk = True
                End If
              
              Else
                 vOk = True
              End If
              
              '...ensure it hasn't been purchased already                
              If vOk And Instr(vEcom_Programs, vProg_Id) = 0 Then
                vOk = True
                Exit For    '...as long as there's one program available, then display the group
              End If              
              
            End If

          End If

         '...if we haven't exitted the loop then assume this program cannot be listed
          vOk = False
  
        Next
 
      
        If vOk Then       
          vCnt = vCnt + 1
          vBg = "" : If vCnt Mod 2 = 0 Then vBg = "bgcolor='#DDEEF9' bordercolor='#FFFFFF'"   '...color ever other line        

          '...generate a script to launch first catalogue item in the right frame unless overwritten by vTraining=1234 (via URL)
          If vCnt = 1 Then
            Response.Write "<script>{parent.frames.Right.location.href='Ecom" & fIf(Session("Ecom_Media") = "Group2" Or Session("Ecom_Media") = "AddOn2", 3, 2) & "Programs.asp?vCatlNo=" & fIf(Request("vInitCatlNo").Count > 0, Request("vInitCatlNo"), vCatl_No) & "';}</script>"
          End If
    %>
    <tr>
      <td <%=vBg%>>
        <a <%=fStatX%> target="Right" href="Ecom<%=fIf(Session("Ecom_Media") = "Group2" Or Session("Ecom_Media") = "AddOn2", 3, 2)%>Programs.asp?vCatlNo=<%=vCatl_No%>"><%=vCatl_Title%></a><%=fPromo(vCatl_Promo)%>
      </td>
    </tr>
    <%
        End If
        oRs2.MoveNext
      Loop
    %>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>
</html>


