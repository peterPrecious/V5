<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->


<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>


  <title>My Content</title>
  <base target="Details">
</head>

<body onload="parent.frames[1].location.href='Ecom2MyModules.asp'" topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" text="#000080" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <table border="0" width="100%" cellpadding="3" style="border-collapse: collapse" id="table3">
    <tr>
      <td nowrap valign="top"><img border="0" src="../Images/Ecom/User1.gif"> </td>
      <td align="center"><p class="c2">
      
        <b>
        <!--webbot bot='PurpleText' PREVIEW='My Programs'--><%=fPhra(000186)%></p>
        <h6 align="center">
        <!--webbot bot='PurpleText' PREVIEW='There are no programs available.'--><%=fPhra(000003)%> </h6>
  
        <% 
          '... let non learners in unless group2 
          sGetCust svCustId

          If svMembLevel > 2 Then 
            If svMembLevel = 5 Or (svMembLevel < 5  And vCust_MaxUsers >= 0) Then         
        %>   
          <h6 align="center">
          As a <%=fIf(svMembLevel=3, "facilitator", fIf(svMembLevel=4, "manager", "administrator"))%> you can click on &quot;All Programs&quot; to see a complete listing. 
          </h6>

          <h2 align="center">
          <a <%=fStatX%> target="_self" href="Ecom2MyPrograms.asp"fPhraH(000186)/a> | 
          <a <%=fStatX%> target="_self" href="Ecom2MyPrograms.asp?vMode=All"fPhraH(000068)/a>
          </h2>        

        <% 
            End If      
          End If 
        %> 
        </b>     
      </td>
    </tr>
  </table>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

  <p class="c2" align="center"><a target="_parent" href="MyContent.asp">.</a></p>

  </body>

</html>



