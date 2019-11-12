<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<html>

<head>
  <title>CatEdit</title>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>

  <style>
    #header tr td { padding: 20px; }
  </style>
</head>

<body>

  <%

  Server.Execute vShellHi
 
  Dim vFunction, vAction, aPrograms, vCnt

  vFunction = ""

  vAction = Request("vAction")

  '...add, delete, edit, clone or change sort order of catalogue
  Select Case Left(vAction, 2)

    Case "AD" '...add an empty group item
      sAddCatl
    Case "DL" '...delete a group item
      vCatl_No = Clng(Mid(vAction, 4))
      sDeleteCatl
      sOrderCatl
    Case "UP", "DN", "TP", "BT" '...shift a group item
      vCatl_No = Clng(Mid(vAction, 4))
      sShiftOrder vCatl_No, Left(vAction, 2)
    Case "ED" '...edit a group
      sExtractCatl
      sUpdateCatl
    Case "CL" '...clone a catalogue
      vCust_Id = Request("vCust_Id")
      sCloneCatl vCust_Id
    Case "CC" '...clear out catalogue
      sClearCatl svCustId
      
  End Select
   
    
  If Len(Request("vHidden")) = 0 Then

  %>

  <script>
  function validate(theForm) {

    if (theForm.vCust_Id.value == "") {
      alert("Please enter a value for the \"Customer Id\" field.");
      theForm.vCust_Id.focus();
      return (false);
    }

    if (theForm.vCust_Id.value.length < 8) {
      alert("Please enter at least 8 characters in the \"Customer Id\" field.");
      theForm.vCust_Id.focus();
      return (false);
    }

    if (theForm.vCust_Id.value.length > 8) {
      alert("Please enter at most 8 characters in the \"Customer Id\" field.");
      theForm.vCust_Id.focus();
      return (false);
    }
    return (true);
  }
  </script>

  <form method="POST" action="CatlEdit.asp" target="_self" onsubmit="return validate(this)">
    <table id="header">
      <tr>
        <td style="width: 33%">
          <h1>Catalogue Editor</h1>
          <p class="c2">The Catalogue is a collection of Programs organized in Groups.&nbsp; To edit a Catalogue Group click the Title below.&nbsp; If the Group is Active it will display a checkmark.</p>
        </td>
        <td style="width: 33%">
          <h1>Clone another Catalogue</h1>
          <p class="c2">You can Copy other Customer&#39;s full Catalogue into this site.</p>
          <p style="text-align: center">
            <b>Customer Id: </b>&nbsp;<!--webbot bot="Validation" s-display-name="Customer Id" b-value-required="TRUE" i-minimum-length="8" i-maximum-length="8" --><input type="text" name="vCust_Id" size="9" maxlength="8">
            <input type="submit" value="Clone" name="bCone" class="button070">
          </p>
        </td>
        <td style="width: 33%">
          <h1>Add new Catalogue Group</h1>
          <p class="c2">Click to Add a new empty Catalogue Group at the bottom of the catalogue.</p>
          <p style="text-align: center">
            &nbsp;<a <%=fstatx%> href="CatlEdit.asp?vAction=AD"><input type="button" onclick="location.href = 'CatlEdit.asp?vAction=AD'" value="Add" name="bAdd" class="button070"></a>
        </td>
      </tr>
    </table>
    <input type="hidden" name="vHidden" value="Y"><input type="hidden" name="vAction" value="CL">
  </form>

  <!--- Catalogue List -->
  <table id="list">
    <%
      '...read Catl
      vCnt = 0
      sGetCatls_Rs svCustId    
      Do While Not oRs2.Eof 
        sReadCatl
        If Len(Trim(fOkValue(vCatl_Title))) = 0 Then vCatl_Title = "N/A"
        vCnt = vCnt + 1
        If vCnt = 1 Then
    %>
    <tr>
      <th class="rowshade" style="width: 30%">Group Title</th>
      <th class="rowshade" style="width: 10%">Active?</th>
      <th class="rowshade" style="width: 10%">Order</th>
      <th class="rowshade" style="width: 40%">Programs</th>
      <th class="rowshade" style="width: 10%">&nbsp;</th>
    </tr>
    <%        
        End If
    %>
    <tr>
      <td><a <%=fstatx%> href="CatlEdit.asp?vEditCatlNo=<%=vCatl_No%>&vHidden=n"><%=fLeft(vCatl_Title, 48)%></a></td>
      <td style="text-align: center"><% =fIf (vCatl_Active, "<img border='0' src='../Images/Common/CheckMark.jpg' width='12' height='15'>", "") %></td>
      <td>
        <a <%=fstatx%> title="Move to top of the list" href="CatlEdit.asp?vAction=TP_<%=vCatl_No%>">
          <img border="0" src="../Images/Icons/ArrowTop.gif" width="18" height="22"></a>
        <a <%=fstatx%> title="Move up one line" href="CatlEdit.asp?vAction=UP_<%=vCatl_No%>">
          <img border="0" src="../Images/Icons/ArrowUp.gif" width="18" height="22"></a>
        <a <%=fstatx%> title="Move down one line" href="CatlEdit.asp?vAction=DN_<%=vCatl_No%>">
          <img border="0" src="../Images/Icons/ArrowDown.gif" width="18" height="22"></a>
        <a <%=fstatx%> title="Move to bottom of the list" href="CatlEdit.asp?vAction=BT_<%=vCatl_No%>">
          <img border="0" src="../Images/Icons/ArrowBottom.gif" width="18" height="22"></a>
      </td>
      <td>
        <p class="c2">
          <%
          If Len(vCatl_Programs) > 0 Then 
            aPrograms = Split(vCatl_Programs, " ")
            For i = 0 to Ubound(aPrograms)
              vProg_Id = Left(aPrograms(i), 7)
          %>
          <a <%=fstatx%> target="_blank" href="Program.asp?vEditProgId=<%=vProg_Id%>&vClose=Y&vHidden=n"><%=vProg_Id%></a>
          <%
  	        Next
  	      Else
  	        Response.Write "&nbsp;&nbsp;"
  	      End If
          %>
        </p>
      </td>
      <td>
        <a <%=fstatx%> href="CatlEdit.asp?vAction=DL_<%=vCatl_No%>"></a>
        <input type="button" onclick="jconfirm('CatlEdit.asp?vAction=DL_<%=vCatl_No%>', '<%=Server.HtmlEncode("Ok to Delete this catalogue item?")%>')" value="Delete" name="bDelete" class="button070">
      </td>
    </tr>
    <%  
        oRs2.MoveNext
      Loop
      Set oRs2 = Nothing
      sCloseDb2    
    %>
  </table>


  <%  If vCnt > 0 Then %>
  <div style="width: 50%; margin: 20px auto; border: 1px solid red; text-align: center; padding: 10px;" class="c6">
    Clear out the entire Catalogue.<br />
    <br />
    Note: This is an irreversible, irrecoverable action !<br />
    <br />
    <input type="button" onclick="jconfirm('CatlEdit.asp?vAction=CC', '<%=Server.HtmlEncode("Ok to clear out catalogue?")%>')" value="Clear" name="bClear" class="button070">
  </div>

  <%  Else %>

  <h2>There are currently no items in the Catalogue. </h2>

  <%
      End If
    
    Else

      If Len(Request.QueryString("vEditCatlNo")) > 0 Then 
        vCatl_No = Request.QueryString("vEditCatlNo")
        sGetCatl (vCatl_No)         
      Else
         Response.Redirect "CatlEdit.asp"          
      End If
  %>

  <form method="POST" action="CatlEdit.asp" target="_self">
    <input type="hidden" name="vAction" value="ED"><input type="hidden" name="vCatl_Order" value="<%=vCatl_Order%>">
    <table class="table">
      <tr>
        <td colspan="2">
          <h1>Catalogue Group</h1>
          <h2>Click <b>Update</b> if you edit any values in this Catelogue Group.</h2>
        </td>
      </tr>
      <tr>
        <th>Group Title : </th>
        <td>
          <input type="text" size="46" name="vCatl_Title" value="<%=vCatl_Title%>" maxlength="500">
        </td>
      </tr>
      <tr>
        <th>Promo :</th>
        <td>
          <input type="text" size="71" name="vCatl_Promo" value="<%=vCatl_Promo%>" maxlength="256"><br>
          Enter any promotional text to appear in More Content, italicized in red below the title as follows:
          <br>
          <font color="#FF0000"><i>Do not enter any HTML tags.</i></font>
        </td>
      </tr>
      <tr>
        <th>Active ? </th>
        <td>
          <input type="radio" value="1" name="vCatl_Active" <%=fcheck(fsqlboolean(vcatl_active), 1)%>>Yes&nbsp; 
          <input type="radio" value="0" name="vCatl_Active" <%=fcheck(fsqlboolean(vcatl_active), 0)%>>No
          <br>
          If inactive, this catalogue item will not be available for purchase but can be accessed if already purchased.
        </td>
      </tr>
      <tr>
        <th>Program Strings : </th>
        <td>
          <textarea rows="6" name="vCatl_Programs" cols="52"><%=vCatl_Programs%></textarea><br>
          Click to access programs :
          <br>
          <%
          If Len(vCatl_Programs) > 0 Then 
            aPrograms = Split(vCatl_Programs, " ")
            For i = 0 to Ubound(aPrograms)
              vProg_Id = Left(aPrograms(i), 7)
          %>
          <a <%=fstatx%> target="_blank" href="ProgramEdit.asp?vEditProgId=<%=vProg_Id%>&vHidden=n&vClose=Y"><%=vProg_Id%></a>
          <%
  	        Next
  	      End If
          %>
          <br>
          &nbsp;
        </td>
      </tr>
      <tr>
        <td colspan="2" style="text-align: center;">&nbsp;<p>
          <input type="button" onclick="javascript: history.back(1)" value="Return" name="bReturn" id="bReturn" class="button070"><%=f10%><input type="submit" value="Update" name="bUpdate" class="button070">
        </p>
          <p>
            &nbsp;
        </td>
      </tr>
    </table>
    <input type="hidden" name="vCatl_No" value="<%=vCatl_No%>">
  </form>

  <%
    End If  
  %>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
