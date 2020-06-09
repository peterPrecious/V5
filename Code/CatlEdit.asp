<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Catl.asp"-->
<!--#include virtual = "V5/Inc/Db_Prog.asp"-->

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
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
  <!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Type="text/javascript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.vCust_Id.value == "")
  {
    alert("Please enter a value for the \"Customer Id\" field.");
    theForm.vCust_Id.focus();
    return (false);
  }

  if (theForm.vCust_Id.value.length < 8)
  {
    alert("Please enter at least 8 characters in the \"Customer Id\" field.");
    theForm.vCust_Id.focus();
    return (false);
  }

  if (theForm.vCust_Id.value.length > 8)
  {
    alert("Please enter at most 8 characters in the \"Customer Id\" field.");
    theForm.vCust_Id.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="CatlEdit.asp" target="_self" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1" language="JavaScript">
    <table border="0" width="100%" cellpadding="10" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td valign="top" width="33%" align="left">
        <h1>Catalogue Editor</h1>
        <h2>The Catalogue is a collection of Program organized in Groups.&nbsp; To edit a Catalogue Group click the Title below.&nbsp; If the Group is Active it will display a checkmark.</h2>
        </td>
        <td valign="top" width="33%" align="left">
        <h1>Clone another Catalogue</h1>
        <h2>You can Copy other Customer&#39;s full Catalogue into this site.</h2>
        <p align="right"><b>Customer Id: </b>&nbsp;<!--webbot bot="Validation" s-display-name="Customer Id" b-value-required="TRUE" i-minimum-length="8" i-maximum-length="8" --><input type="text" name="vCust_Id" size="9" maxlength="8"> <input type="submit" value="Clone" name="bCone" class="button070"></td>
        <td align="left" width="33%" valign="top">
        <h1>Add new Catalogue Group</h1>
        <h2><span style="font-weight: 400">Click to Add a new empty Catalogue Group at the bottom of the catalogue.</span></h2><p align="right">&nbsp;<a <%=fstatx%> href="CatlEdit.asp?vAction=AD"><input type="button" onclick="location.href='CatlEdit.asp?vAction=AD'" value="Add" name="bAdd" class="button070"></a></td>
      </tr>
      <input type="hidden" name="vHidden" value="Y"><input type="hidden" name="vAction" value="CL">
    </table>
  </form>
  <!--- Catalogue List -->
  <table width="100%" border="1" cellpadding="3" cellspacing="0" bordercolor="#DDEEF9" style="border-collapse: collapse">
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
      <th align="left" nowrap bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Group Title</th>
      <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Active?</th>
      <th bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Order</th>
      <th align="left" bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">Programs</th>
      <th align="right" bgcolor="#DDEEF9" bordercolor="#FFFFFF" height="30">&nbsp;</th>
    </tr>
    <%        
        End If
    %>
    <tr>
      <td valign="top" nowrap><a <%=fstatx%> href="CatlEdit.asp?vEditCatlNo=<%=vCatl_No%>&vHidden=n"><%=fLeft(vCatl_Title, 48)%></a></td>
      <td valign="top" align="center" nowrap><% =fIf (vCatl_Active, "<img border='0' src='../Images/Common/CheckMark.jpg' width='12' height='15'>", "") %></td>
      <td valign="top" align="center" nowrap>&nbsp; <a <%=fstatx%> title="Move to top of the list" href="CatlEdit.asp?vAction=TP_<%=vCatl_No%>"><img border="0" src="../Images/Icons/ArrowTop.gif" width="18" height="22" longdesc="Move up to top"></a><a <%=fstatx%> title="Move up one line" href="CatlEdit.asp?vAction=UP_<%=vCatl_No%>"><img border="0" src="../Images/Icons/ArrowUp.gif" width="18" height="22" longdesc="Move up one group"></a><a <%=fstatx%> title="Move down one line" href="CatlEdit.asp?vAction=DN_<%=vCatl_No%>"><img border="0" src="../Images/Icons/ArrowDown.gif" width="18" height="22" longdesc="Move down one group"></a><a <%=fstatx%> title="Move to bottom of the list" href="CatlEdit.asp?vAction=BT_<%=vCatl_No%>"><img border="0" src="../Images/Icons/ArrowBottom.gif" width="18" height="22" longdesc="Move to the bottom"></a>&nbsp;&nbsp;&nbsp; </td>
      <td valign="top">
      <p class="c2">
      <%
          If Len(vCatl_Programs) > 0 Then 
            aPrograms = Split(vCatl_Programs, " ")
            For i = 0 to Ubound(aPrograms)
              vProg_Id = Left(aPrograms(i), 7)
      %> 
      <a <%=fstatx%> target="_blank" href="ProgramEdit.asp?vEditProgId=<%=vProg_Id%>&vClose=Y&vHidden=n"><%=vProg_Id%></a> 
      <%
  	        Next
  	      Else
  	        Response.Write "&nbsp;&nbsp;"
  	      End If
      %> </td>
      <td valign="top" align="right"><a <%=fstatx%> href="CatlEdit.asp?vAction=DL_<%=vCatl_No%>"><input type="button" onclick="location.href='CatlEdit.asp?vAction=DL_<%=vCatl_No%>'"  value="Delete" name="bDelete" class="button070"></a></td>
    </tr>
    <%  
        oRs2.MoveNext
      Loop
      Set oRs2 = Nothing
      sCloseDb2    
    %>
  </table>

  <%  If vCnt > 0 Then %> 
    <p align="center" class="c2">Click to Clear out the entire Catalogue.</p>
    <p align="center" class="c5">Note this is an irreversible, irrecoverable action !</p>
    <p align="center"><a <%=fstatx%> href="javascript:jconfirm('CatlEdit.asp?vAction=CC','<%=Server.HtmlEncode(fPhraH(000303))%>')"><img border="0" src="../Images/Buttons/Clear_<%=svLang%>.gif"></a></p>

  <%  Else %>
    <h1 align="center">There are currently no items in the Catalogue. </h1>

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
    <table border="1" width="100%" cellspacing="0" cellpadding="3" bordercolor="#DDEEF9" style="border-collapse: collapse">
      <tr>
        <td align="center" width="100%" valign="top" colspan="2">
        <h1>Catalogue Group</h1>
        <h2>Click <b>Update</b> if you edit any values in this group.</h2>
        </td>
      </tr>
      <tr>
        <th align="right" width="30%" valign="top" nowrap>Group Title : </th>
        <td width="70%"><input type="text" size="46" name="vCatl_Title" value="<%=vCatl_Title%>" maxlength="500"></td>
      </tr>
      <tr>
        <th align="right" width="35%" valign="top" nowrap>Promo :</th>
        <td width="65%" valign="top"><input type="text" size="71" name="vCatl_Promo" value="<%=vCatl_Promo%>" maxlength="256" class="c2"><br>Enter any promotional text to appear in More Content, italicized in red below the title as follows: <br><font color="#FF0000"><i>Do not enter any HTML tags.</i></font></td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap width="30%">Active ? </th>
        <td valign="top" width="70%">
          <input type="radio" value="1" name="vCatl_Active" <%=fcheck(fsqlboolean(vcatl_active), 1)%>>Yes&nbsp; 
          <input type="radio" value="0" name="vCatl_Active" <%=fcheck(fsqlboolean(vcatl_active), 0)%>>No <br>
          If inactive, this catalogue item will not be available for purchase but can be accessed if already purchased.
        </td>
      </tr>
      <tr>
        <th align="right" valign="top" nowrap width="30%">Program Strings : </th>
        <td valign="top" width="70%"><textarea rows="6" name="vCatl_Programs" cols="52"><%=vCatl_Programs%></textarea><br>Click to access programs : <br>
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
        <br>&nbsp; </td>
      </tr>
      <tr>
        <td align="center" width="100%" valign="top" colspan="2" height="38">&nbsp;<p><input type="button" onclick="javascript:history.back(1)" value="Return" name="bReturn" id="bReturn" class="button070"><%=f10%><input type="submit" value="Update" name="bUpdate" class="button070"></p>
        <p>&nbsp;</td>
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

