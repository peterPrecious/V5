<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->
<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Document.asp"-->

<% 
  sGetCust (svCustId) 
  sGetMemb (svMembNo)    
  Dim bCompletion : bCompletion = fIf((svMembManager OR svMembLevel = 5) AND Instr("CNPX HMVC INDG SAPU CAST UGRC", Left(svCustId, 4)) > 0, True, False) 
%>

<html>

<head>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js"></script>
  <script src="/V5/Inc/Functions.js"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js"></script><% End If %>
  <link href="/V5/Inc/Vubiz.css" type="text/css" rel="stylesheet">
  <style>
    .notice, ul, li { color: red; text-align: center; }
      .notice li { text-align: left; }
      .notice a { font-weight: normal; }
  </style>
  <script>
    function docWindow(url) {
      var docs = window.open(url,'Document','toolbar=no,width=600,height=800,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    }
  </script>
</head>

<body leftmargin="0" topmargin="0" bgcolor="#FFFFFF" text="#000000" link="#000080" vlink="#000080" alink="#000080">

  <% Server.Execute vShellHi %>
  <div align="center">
    <table class="tableBorder" id="table2" cellspacing="0" cellpadding="10" style="width:600px; text-align:center">
      <tr>
        <td height="0" width="0">
        <table border="0" cellpadding="5" id="table3" style="width:100%;">
          <!------------------------------------------------------- Facilitator -->
          <% If svMembLevel > 2 Then %>
          <tr>
            <th>
            <h1>
            <!--webbot bot='PurpleText' PREVIEW='Learning Management System'--><%=fPhra(000167)%></h1>
            </th>
          </tr>
          <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
            <th class="underline overline" height="25" align="left">
            <!--webbot bot='PurpleText' PREVIEW='Facilitator Services'--><%=fPhra(000259)%> </th>
          </tr>
          <tr>
            <td nowrap>
            <table cellpadding="2" width="100%" border="0" bordercolor="#FFFFFF">
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="User<%=fGroup%>.asp">
                <!--webbot bot='PurpleText' PREVIEW='My Profile'--><%=fPhra(000185)%></a></td>
                <td align="left">&nbsp;</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td >&nbsp; <a <%=fstatx%> href="User<%=fGroup%>.asp?vMembNo=0">
                <!--webbot bot='PurpleText' PREVIEW='Add a Learner'--><%=fPhra(000370)%></a></td>
                <td  align="left"></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="Users.asp">
                <!--webbot bot='PurpleText' PREVIEW='Learner Report'--><%=fPhra(000367)%></a></td>
                <td align="left">&nbsp;</td>
              </tr>

              <!-- old reports used by corporate - to be transferred over in time -->
              <%  If (vCust_Level = 4 OR svMembLevel = 5) Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp;</td>
                <td class="notice" style="text-align: left">&nbsp;</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="LearnerReportCard.asp">
                <!--webbot bot='PurpleText' PREVIEW='Learner Report Card'--><%=fPhra(000795)%></a></td>
                <td class="notice" style="text-align: left">&nbsp;</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="Activity.asp">
                <!--webbot bot='PurpleText' PREVIEW='Activity Report'--><%=fPhra(000487)%></a></td>
                <td class="notice" style="text-align: left">&nbsp;</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="LogReport5.asp">
                <!--webbot bot='PurpleText' PREVIEW='Assessment Report'--><%=fPhra(000074)%></a>&nbsp; </td>
                <td class="notice" style="text-align: left">&nbsp;</td>
              </tr>
              <%   If fIsCorporate Or svMembLevel > 4 Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="LogReport4.asp">
                <!--webbot bot='PurpleText' PREVIEW='Completion Report - Basic'--><%=fPhra(001313)%></a></td>
                <td class="notice" style="text-align: left">&nbsp;</td>
              </tr>
              <%   End If %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp;</td>
                <td class="notice" style="text-align: left">&nbsp;</td>
              </tr>
              <% End If %>


              <!-- new reports used by channels  -->
              <%  If (vCust_Level < 3 Or svMembLevel = 5) Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="RTE_History.asp">
                <!--webbot bot='PurpleText' PREVIEW='Learner Report Card'--><%=fPhra(000795)%></a></td>
                <td class="notice" style="text-align: left">&nbsp;</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="/Gold/vuClientReportingDev/AssReportFilter.aspx?AccountID=<%=svCustAcctId%>&MembNo=<%=svMembNo%>&reportId=1">
                <!--webbot bot='PurpleText' PREVIEW='Activity Report'--><%=fPhra(000487)%></a></td>
                <td class="notice" style="text-align: left">&nbsp;</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="/Gold/vuClientReportingDev/AssReportFilter.aspx?AccountID=<%=svCustAcctId%>&MembNo=<%=svMembNo%>&reportId=2">
                <!--webbot bot='PurpleText' PREVIEW='Assessment Report'--><%=fPhra(000074)%></a></td>
                <td class="notice" style="text-align: left">&nbsp;</td>
              </tr>
              <%   If fIsCorporate Or svMembLevel > 4 Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="/Gold/vuClientReportingDev/AssReportFilter.aspx?AccountID=<%=svCustAcctId%>&MembNo=<%=svMembNo%>&reportId=3">
                <!--webbot bot='PurpleText' PREVIEW='Completion Report - Basic'--><%=fPhra(001313)%></a></td>
                <td class="notice" style="text-align: left">&nbsp;</td>
              </tr>
              <%   End If %> 

              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp;</td>
                <td class="notice" style="text-align: left">&nbsp;</td>
              </tr>

              <% End If %>
              

              
              <% If vCust_Id = "CCHS2074" Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="/V5/Repository/V5_Vubz/8108/Tools/CPR_CCHS2074.asp">CPR Learner Activity Report</a></td>
                <td align="left">&nbsp;</td>
              </tr>
              <% End If %> 
							
							
							<% If fIsGroup2 Or svMembLevel = 5 Then %>

              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="ProgramsAssigned.asp"><!--webbot bot='PurpleText' PREVIEW='Programs Purchased and Assigned'--><%=fPhra(000886)%></a></td>
                <td align="left">&nbsp;</td>
              </tr>


              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> target="_blank" href="/Gold/vuClientReporting/ReportViewerFrame.aspx?AccountID=<%=svCustAcctId%>&reportfile=App_Data/repLearnerCompletion.frx"><!--webbot bot='PurpleText' PREVIEW='Completion Status Report - Online Version'--><%=fPhra(001588)%></a></td>
                <td class="notice" style="text-align: left; width:35%"><a class="green" href="javascript:toggle('div_R1');">Description</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <th nowrap colspan="2" valign="bottom">
                <div style="text-align: left; margin-left:20px" id="div_R1" class="div">
                  <table border="0" id="table3" cellpadding="10" style="border-collapse: collapse" bordercolor="#DDEEF9" bgcolor="#FFFFFF">
                    <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                      <td><!--webbot bot='PurpleText' PREVIEW='Lists completions/incompletions for all<br />programs assigned to your Learners.'--><%=fPhra(001589)%></td>
                    </tr>
                  </table>
                </div>
                </th>
              </tr>
              <tr>
                <td>&nbsp; <!--webbot bot='PurpleText' PREVIEW='Completion Status Report - CSV Download'--><%=fPhra(001590)%></td>
                <td class="notice" style="text-align: left; width:35%">...being revised...</td>
              </tr>
              <tr>
                <th nowrap colspan="2" valign="bottom">
                <div style="text-align: left; margin-left:20px" id="div_R2" class="div">
                  <table border="0" id="table3" cellpadding="10" style="border-collapse: collapse" bordercolor="#DDEEF9" bgcolor="#FFFFFF">
                    <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                      <td><!--webbot bot='PurpleText' PREVIEW='Creates a CSV file of completions/incompletions<br />for all programs assigned to your Learners'--><%=fPhra(001591)%></td>
                    </tr>
                  </table>
                </div>
                </th>
              </tr>

              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> target="_blank" href="/Gold/vuclientreporting/ReportExport.aspx?AccountID=<%=svCustAcctId%>&reportfile=repLearnerIncompleteCourseExport.frx&type=CSV"><!--webbot bot='PurpleText' PREVIEW='InCompletion Report - CSV Download'--><%=fPhra(001592)%></a> </td>
                <td class="notice" style="text-align: left; width:35%"><a class="green" href="javascript:toggle('div_R3');">Description</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <th nowrap colspan="2" valign="bottom">
                <div style="text-align: left; margin-left:20px" id="div_R3" class="div">
                  <table border="0" id="table3" cellpadding="10" style="border-collapse: collapse" bordercolor="#DDEEF9" bgcolor="#FFFFFF">
                    <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                      <td style="text-align: left; margin-left: 10px"><!--webbot bot='PurpleText' PREVIEW='Creates a CSV file of programs that have not been completed<br />for all programs assigned to your Learners'--><%=fPhra(001774)%></td>
                    </tr>
                  </table>
                </div>
                </th>
              </tr>

              <%   If svMembLevel = 5 Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td style="vertical-align: top">&nbsp;&nbsp;<!--webbot bot='PurpleText' PREVIEW='My Custom Policies'--><%=fPhra(001558)%></td>
                <td class="notice" style="text-align: left; width:35%"><a class="green" href="javascript:toggle('div_R4');">Description</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">
                  <div style="text-align: left; margin-left:20px" id="div_R4" class="div">
                    <table border="0" id="table1" cellpadding="10" style="border-collapse: collapse" bordercolor="#DDEEF9" bgcolor="#FFFFFF">
                      <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                        <td>These are the policies that will be rendered from the<br />&quot;smartLinks&quot; in your content from this Account.<br /><br />&nbsp;&nbsp;&nbsp; 
                          <a href="#" onclick="docWindow('<%=fDocumentUrl("Harassment.pdf", "", svLang, Left(svCustId, 4), svCustAcctId, "", "")%>')">Harassment.pdf</a><br />&nbsp;&nbsp;&nbsp; 
                          <a href="#" onclick="docWindow('<%=fDocumentUrl("Affirmative.pdf", "", svLang, Left(svCustId, 4), svCustAcctId, "", "")%>')">Affirmative.pdf</a><br />&nbsp;&nbsp;&nbsp; 
                          <a href="#" onclick="docWindow('<%=fDocumentUrl("Conflict.pdf", "", svLang, Left(svCustId, 4), svCustAcctId, "", "")%>')">Conflict.pdf</a>
                        </td>
                      </tr>
                    </table>
                  </div>
                </td>
              </tr>
              <%   End If %> 
              <% End If %>
            </table>
            </td>
          </tr>
          <% End If %>
          <!------------------------------------------------------- Manager --><% If svMembLevel > 3 Then %>
          <tr id="t2" class="c1">
            <th class="underline" height="25" align="left">&nbsp;</th>
          </tr>
          <tr>
            <th class="underline" height="25" align="left">Advanced Services </th>
          </tr>
          <tr>
            <td>
            <table cellpadding="2" width="100%" border="0" bordercolor="#FFFFFF">
              <!--                <tr onMouseOver="this.className='bgOn'" onMouseOut="this.className='bgOff'"><td>&nbsp; <a <%=fstatx%> href="/V5/Repository/Documents/History/Default.asp">Assessment Data Dump</a></td></tr>-->
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="EcomReport.asp">Ecommerce Report - Basic</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="EcomReport0.asp">Ecommerce Report - Advanced</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="EcomReport1.asp">Ecommerce Sales Summary Report</a></td>
              </tr>
              <%    If vMemb_Ecom Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="Ecom2Start.asp?vMode=More&vContentOptions=<%=vContentOptions%>">Post Manual Ecommerce Sales</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="EcomExtend.asp">Extend Access for Ecommerce Purchase</a></td>
              </tr>
              <%    End If %> <%    If fIsCorporate Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="AccessReport.asp">Programs Assigned Report - Corporate Sites</a></td>
              </tr>
              <%    End If %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="EcomReport3.asp">Program Sales Report</a></td>
              </tr>
              <%    If vCust_Id = "CAAM3001" Or svMembLevel = 5 Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="EcomReport4.asp">Ecommerce Completion Report</a></td>
                <td class="notice" width="35%" style="text-align: left">New</td>
              </tr>
              <%    End If %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="Catalogue.asp">Catalogue Summary</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="CatlByCustId_x.asp">Catalogue Nos (Excel)</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="CustomerExpires.asp">Maintain Customer Expiry Date</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="CustomerExpiryReport.asp">Customer Expiry Report</a></td>
              </tr>
              <%    If svCustIssueIds Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="IssueIds.asp">Generate Multiple Passwords</a></td>
              </tr>
              <%    End If %> <%    If svMembLevel = 5 And Not fNoValue(vCust_IssueIdsTemplate) Then %> <%    End If %> <%  If svMembLevel = 5 Or (svMembLevel = 4 And vMemb_MyWorld) Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="TaskEdit1.asp">Maintain Tasks for My Learning</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="CritEdit.asp">Maintain Group 1 Table</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="JobsEdit.asp">Maintain Jobs Table</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="SkilEdit.asp">Maintain Skills Table</a></td>
              </tr>
              <%  End If %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="UsersBulkInput.asp">Upload Learners (Basic - for smaller accounts)</a></td>
              </tr>
              <%  If svMembLevel > 3 Then %> <%    If (Instr("CMSS2592 UGRC1464", svCustId) > 0) Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="<%=svDomain%>/Repository/Import/<%=svCustId%>.asp">Upload Learners (Custom)</a></td>
              </tr>
              <%    Else %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="<%=svDomain%>/Repository/Upload/Upload2.asp">Upload Learners (Advanced - for larger accounts)</a></td>
              </tr>
              <%    End If %> <%  End If %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="DocumentUpload.asp">Upload Custom Document</a> (for content smartlinks)</td>
              </tr>
              <!--                 </div>-->
            </table>
            </td>
          </tr>
          <% End If %>
          <!------------------------------------------------------- Customer Management System--><% If svMembLevel = 5 Or (svMembLevel = 4 And vMemb_VuBuild) Then %>
          <tr>
            <th class="underline" height="25" align="left">&nbsp;</th>
          </tr>
          <tr id="t4" class="c1">
            <th class="underline" height="25" align="left">Customer Management</th>
          </tr>
          <tr>
            <td>
            <table cellpadding="2" width="100%" border="0" bordercolor="#FFFFFF">
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="CatlEdit.asp">Customer Catalogue</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
<!--            <td>&nbsp; <a <%=fstatx%> href="CustomerEdit.asp?vEditCustId=<%=svCustId%>&vHidden=n">Customer Profile</a> </td>-->
                <td>&nbsp; <a <%=fstatx%> href="Customer.asp?vEditCustId=<%=svCustId%>&vHidden=n">Customer Profile</a> </td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="CustomerList.asp">Customer List</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="CustomerNotes.asp">Customer Notes</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="CatlFind.asp">Find Programs or Modules in Catalogues</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="CustomerActivityReport.asp">Customer Activity Report</a></td>
              </tr>
              <%   If svMembLevel = 5 Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="CustomerSellerOwners.asp">Customer Sellers | Owners Report</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="LogsHistoryTransfer.asp">Transfer Learner History</a></td>
              </tr>
              <%   End If %>
            </table>
            </td>
          </tr>
          <% End If %>
          <!------------------------------------------------------- Content Management System --><% If svMembLevel = 5 Or vMemb_LCMS Then %>
          <tr class="underline" id="t5" class="c1">
            <th class="underline" height="25" align="left">&nbsp;</th>
          </tr>
          <tr class="underline" id="t5" class="c1">
            <th class="underline" height="25" align="left">Content Management</th>
          </tr>
          <tr>
            <td>
            <table cellspacing="0" cellpadding="2" width="100%">
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="ProgramEdit.asp">Maintain Programs</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="Modules.asp">Maintain Modules</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="ModSkillSet.asp">Update Module Skill Sets</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="javascript:jconfirm('ProgramUpdate.asp','OK to Continue?  This job can takes a long time to complete.')">Update Program Lengths</a></td>
              </tr>
            </table>
            </td>
          </tr>
          <% End If %>
          <!------------------------------------------------------- Vubiz Internal--><% If svMembLevel = 5 Then %>
          <tr id="t6" class="c1">
            <th class="underline" height="25" align="left">&nbsp;</th>
          </tr>
          <tr>
            <th class="underline" height="25" align="left">Vubiz Internal</th>
          </tr>
          <tr>
            <td>
            <table cellpadding="2" width="100%" border="0" bordercolor="#FFFFFF">
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="RTE_CreateSession.asp">Create an RTE Session</a></td>
                <td class="notice" width="35%" style="text-align: left">Ferret management</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="Channel_Management.asp">Channel Analysis</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="CertificateSales.asp">Certificate Sales Report</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="/V5/Repository/Documents/MultiUserManual/UploadMultiUserManual.asp">Upload Multi User Manuals</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="/V5/RTEResetAttempts.htm">Reset Attempts on Scorm RTE (Temp)</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="/V5/Repository/V5_Vubz/0000/Tools/ERGP/Customers.asp">ERGP - Customer Control Center</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="EcomHistory.asp">Ecommerce Transactions NOT Processed</a>&nbsp; </td>
                <td class="notice" width="35%" style="text-align: left">Under review</td>
              </tr>
              <% If Len(Trim(vCust_ParentId)) = 0 Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="SnapshotExam.asp">Course Completion Snapshot</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="Patience.asp?vNext=SnapshotAll.asp">Course Usage Snapshot</a></td>
              </tr>
              <% End If %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a href="EcomMemo.asp">Post Memo Funds to InternetSecure</a></td>
              </tr>
              <% If svMembInternal Then %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="UserAdmin.asp">Add Super Administrators</a></td>
              </tr>
              <% End If %>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a href="../vuNews.asp">vuNews Email Addresses</a></td>
                <td class="notice" width="35%" style="text-align: left">Excel for Jen</td>
              </tr>
            </table>
            </td>
          </tr>
          <% End If %>
          <!------------------------------------------------------- Client Review --><% If svMembLevel > 3 Then %> <%   If Left(svCustId, 4) = "VUBZ" Or Left(svCustId, 4) = "ERGP"  Or Left(svCustId, 4) = "CCHS" Then %>
          <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
            <th class="underline" height="25" align="left">&nbsp;</th>
          </tr>
          <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
            <th class="underline" height="25" align="left">Client Review</th>
          </tr>
          <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
            <td>
            <table cellspacing="0" cellpadding="2" width="100%">
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; There are no new features to Review</td>
              </tr>
            </table>
            </td>
          </tr>
          <%   End If %> <% End If %>
          <!------------------------------------------------------- Completion Testing --><% If bCompletion Then %>
          <tr class="underline" id="t8">
            <th height="25" align="left">&nbsp;</th>
          </tr>
          <tr class="underline" id="t8">
            <th height="25" align="left">Completion System</th>
          </tr>
          <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
            <td>
            <table cellpadding="2" width="100%" border="0" bordercolor="#FFFFFF">
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="Completion.asp">Completion Reports</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="Completion_Learners.asp">Learner Report</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="Completion_LocationManager.asp">Location Manager</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="Completion_LearnerLocation.asp?vMemb_No=<%=svMembNo%>">Learner Location</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="Completion_RoleManager.asp">Role Manager</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td colspan="2">&nbsp; <a <%=fstatx%> href="Completion_AddScores.asp">Add Scores</a></td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="Cineplex_Learner.asp">Add a Learner</a></td>
                <td class="notice" width="35%" style="text-align: left">CNPX</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="HMVC%20Pooh/HMVC_Learner.asp">Add a Learner</a></td>
                <td class="notice" width="35%" style="text-align: left">HMVC</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="CAST_Learner.asp">Add a Learner</a></td>
                <td class="notice" width="35%" style="text-align: left">CAST</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="Cineplex_Learners.asp">Learner Report</a></td>
                <td class="notice" width="35%" style="text-align: left">CNPX</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="HMVC%20Pooh/HMVC_Learners.asp">Learner Report</a></td>
                <td class="notice" width="35%" style="text-align: left">HMVC</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="CAST_Learners.asp">Learner Report</a></td>
                <td class="notice" width="35%" style="text-align: left">CAST</td>
              </tr>
              <tr onmouseover="this.className='bgOn'" onmouseout="this.className='bgOff'">
                <td>&nbsp; <a <%=fstatx%> href="Completion_ReportSet.asp">Advanced Reports</a></td>
                <td class="notice" width="35%" style="text-align: left">Bryan</td>
              </tr>
            </table>
            </td>
          </tr>
          <% End If %>
        </table>
        </td>
      </tr>
    </table>
  </div>
  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>


