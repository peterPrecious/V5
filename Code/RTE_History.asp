<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<%
  Dim vCurN : vCurN = fDefault(Request("vCurN"), 0)
  Dim vNext : vNext = fIf(vCurN = 0, "RTE_History_F.asp", "RTE_History_O.asp")

  'stop 

%>

<html>
  <head>
    <title>RTE_History</title>
    <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
    <script src="/V5/Inc/jQuery.js"></script>
    <script src="/V5/Inc/Functions.js"></script>
    <script src="/V5/Inc/jQueryC.js"></script>
    <script>
      // grab any parms from cookies and pass into the app
      $(document).ready ( 
        function () {
//        debugger;
          if ($.cookie("History_<%=svCustId%>_vSave") == "y") {
            var url    = "<%=vNext%>"                             +
              "?vCurN=" + <%=vCurN%>                              +
              "&vActv=" + $.cookie("History_<%=svCustId%>_vActv") +
              "&vOutp=" + $.cookie("History_<%=svCustId%>_vOutp") +
              "&vStrD=" + $.cookie("History_<%=svCustId%>_vStrD") +
              "&vEndD=" + $.cookie("History_<%=svCustId%>_vEndD") +
              "&vProg=" + $.cookie("History_<%=svCustId%>_vProg") +
              "&vMods=" + $.cookie("History_<%=svCustId%>_vMods") +
              "&vAssg=" + $.cookie("History_<%=svCustId%>_vAssg") +
              "&vPass=" + $.cookie("History_<%=svCustId%>_vPass") +
              "&vLNam=" + $.cookie("History_<%=svCustId%>_vLNam") +
              "&vMemo=" + $.cookie("History_<%=svCustId%>_vMemo") +
              "&vGrou=" + $.cookie("History_<%=svCustId%>_vGrou") +
              "&vSave=" + $.cookie("History_<%=svCustId%>_vSave") 
          } else {
            var url    = "<%=vNext%>"                             +
              "?vCurN=" + <%=vCurN%>                              +
              "&vActv="                                           +
              "&vOutp="                                           +
              "&vStrD="                                           +
              "&vEndD="                                           +
              "&vProg="                                           +
              "&vMods="                                           +
              "&vAssg="                                           +
              "&vPass="                                           +
              "&vLNam="                                           +
              "&vMemo="                                           +
              "&vGrou="                                           +
              "&vSave=" + $.cookie("History_<%=svCustId%>_vSave") 
        }
        url = url.replace(/null/g, "");  
//      document.write (url); 
//      debugger;
        location.href = url;  
      })       
    </script>
  </head>  
  <body></body>
</html>


