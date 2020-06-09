<!--#include virtual = "V5/Inc/Setup.asp"-->
<!--#include virtual = "V5/Inc/Initialize.asp"-->

<!--#include virtual = "V5/Inc/Db_Phra.asp"-->
<!--#include virtual = "V5/Inc/Db_Cust.asp"-->
<!--#include virtual = "V5/Inc/Db_Memb.asp"-->
<!--#include virtual = "V5/Inc/Document.asp"-->
<!--#include virtual = "V5/Inc/Base64.asp"-->

<% 
  sGetCust (svCustId) 
  sGetMemb (svMembNo)    
  Dim bCompletion : bCompletion = fIf((svMembManager OR svMembLevel = 5) AND Instr("CNPX HMVC INDG SAPU CAST UGRC", Left(svCustId, 4)) > 0, True, False) 
  Dim parms, url

 '...this will put either support@vubiz.com or the customers email address on the Contact Us link at the bottom
  Function fContactUs
    Dim vEmail, vText
    If svCustEmail = "none" Then
      fContactUs = ""
    Else
      vEmail = fIf(Len(svCustEmail) > 0, svCustEmail, "support@vubiz.com")
      Select Case svLang
        Case "FR" : vText = "Communiquez avec nous"
        Case "ES" : vText = "P&#243;ngase en contacto con nosotros"
        Case Else : vText = "Contact Us"
      End Select
'     fContactUs = "<a href='mailto:" & vEmail & "?subject=" & svCustId & " Issue'>" & vText & " (" & vEmail & ")</a>"
      fContactUs = "<a href='mailto:" & vEmail & "?subject=" & svCustId & " Issue'>" & vEmail & "</a>"
    End If
  End Function


%>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html>

<head>
  <title>::Menu</title>
  <meta charset="UTF-8">
  <script src="/V5/Inc/jQuery.js" type="text/javascript"></script>
  <script src="/V5/Inc/Functions.js" type="text/javascript"></script>
  <% If vRightClickOff Then %><script src="/V5/Inc/RightClick.js" type="text/javascript"></script><% End If %>
  <link href="/V5/Inc/Vubi2.css" type="text/css" rel="stylesheet" />
  <script type="text/javascript">
    function docWindow(url) {
      var docs = window.open(url,'Document','toolbar=no,width=600,height=800,left=10,top=10,status=no,scrollbars=no,resizable=yes')
    }
  </script>
  <style type="text/css">
    .table tr:nth-child(even) { background-color: #DDEEF9; }

    .table td:nth-child(1) { width: 60%; }
    .table td:nth-child(2) { width: 40%; }
  </style>
</head>

<body>

  <% Server.Execute vShellHi %>

  <h1>
    <!--[[-->Learning Management System<!--]]--></h1>

  <div style="width: 500px; margin: 25px auto; border: 1px solid navy; padding: 10px 20px 20px 20px;">

    <% If svLang = "EN" Then %>
    <p class="c2">::&nbsp; Help using this service</p>
    &ensp;&ensp;&ensp;&ensp;<a target="_blank" href="../Public/21_FAQ.asp">Click here</a>&nbsp;if you have any questions about how this service works.
    <% End If %>

    <% If svLang = "FR" Then %>
    <p class="c2">::&nbsp; Problèmes liés aux navigateurs?</p>
    &ensp;&ensp;&ensp;&ensp;<a target="_blank" href="../Public/BrowserIssues_FR.htm">Cliquez ici</a> pour options de réglage de votre navigateur Web
    <% End If %>

    <% If svLang = "FR" Then %>
    <p class="c2">::&nbsp; Communiquez avec nous</p>
    &ensp;&ensp;&ensp;&ensp;<span style="text-align: center; margin-top: 20px;"><%=fContactUs%></span>
    <% Else %>
    <p class="c2">::&nbsp; Contact Us</p>
    &ensp;&ensp;&ensp;&ensp;<span style="text-align: center; margin-top: 20px;"><%=fContactUs%></span>
    <% End If%>

    <% If svMembLevel > 3 Then     '...get a temp guid for Portal
      Dim membGuid : membGuid = sp5getMembGuidTemp (vMemb_Guid)    
    %>
    <p class="c2">::&nbsp; New Admin Portal coming soon! Check it out...</p>
    &ensp;&ensp;&ensp;&ensp;<a target="_top" href="/Portal/v7/default.aspx?membGuid=<%=membGuid%>&source=v5">Click here</a> (Please report issues to support@vubiz.com).
    <% End If%>
  </div>


  <table id="menuTable" style="width: 80%; max-width: 800px; text-align: center; margin: auto;">

    <!------------------------------------------------------- Facilitator -->
    <% If svMembLevel > 2 Then %>
    <tr>
      <td class="c2" colspan="2">
        <!--[[-->Facilitator Services<!--]]--></td>
    </tr>
    <tr>
      <td colspan="2">
        <table class="table">
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="User<%=fGroup%>.asp">
              <!--[[-->My Profile<!--]]--></a></td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="User<%=fGroup%>.asp?vMembNo=0">
              <!--[[-->Add a Learner<!--]]--></a></td>
            <td></td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="Users.asp">
              <!--[[-->Learner Report<!--]]--></a></td>
            <td>&nbsp;</td>
          </tr>

          <!-- legacy corporate reports -->
          <%  If (svMembLevel = 5) Then %>
<!--          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="LearnerReportCard.asp">Learner Report Card</a></td>
            <td>Legacy</td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="Activity.asp">Activity Report</a></td>
            <td>Legacy</td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="LogReport5.asp">Assessment Report</a>&nbsp;</td>
            <td>Legacy</td>

          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="LogReport4.asp">Completion Report Basic</a></td>
            <td>Legacy</td>

          </tr>
          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>-->
          <% End If %>


          <!-- new reports used by all with corporate note -->
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="RTE_History.asp">
              <!--[[-->Learner Report Card<!--]]--></a></td>
            <td><% If (fIsCorporate) Then Response.Write "New" %>&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="/Gold/vuClientReportingDev/AssReportFilter.aspx?AccountID=<%=svCustAcctId%>&MembNo=<%=svMembNo%>&reportId=1&vLang=<%=svLang%>">
              <!--[[-->Activity Report<!--]]--></a></td>
            <td><% If (fIsCorporate) Then Response.Write "New" %>&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="/Gold/vuClientReportingDev/AssReportFilter.aspx?AccountID=<%=svCustAcctId%>&MembNo=<%=svMembNo%>&reportId=2&vLang=<%=svLang%>">
              <!--[[-->Assessment Report<!--]]--></a></td>
            <td><% If (fIsCorporate) Then Response.Write "New" %>&nbsp;</td>
          </tr>

          <% If fIsCorporate Then %>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="/Gold/vuClientReportingDev/AssReportFilter.aspx?AccountID=<%=svCustAcctId%>&MembNo=<%=svMembNo%>&reportId=3&vLang=<%=svLang%>"><!--[[-->Completion Report Basic<!--]]--></a></td>
            <td></td>
          </tr>
          <% End If %>

          <tr>
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>          
          
          <% If Left(svCustId, 4) = "ERGP" Or Left(svCustId, 4) = "EVHR" Then %>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="/Gold/vuClientReportingDev/CertReportFilter.aspx?AccountID=<%=svCustAcctId%>&MembNo=<%=svMembNo%>&reportId=3&vLang=<%=svLang%>">Certificate Report (PDF)</a></td>
            <td></td>
          </tr>
          <% End If %>



          <% If vCust_Id = "CCHS2074" Then %>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="/V5/Repository/V5_Vubz/8108/Tools/CPR_CCHS2074.asp">CPR Learner Activity Report</a></td>
            <td>&nbsp;</td>
          </tr>
          <% End If %>


          <% If fIsGroup2 Or svMembLevel = 5 Then %>

          <tr>
            <td>&nbsp; <a <%=fstatx%> href="ProgramsAssigned.asp">
              <!--[[-->Programs Purchased and Assigned<!--]]--></a></td>
            <td>&nbsp;</td>
          </tr>


          <tr>
            <td>&nbsp; <a <%=fstatx%> target="_blank" href="/Gold/vuClientReporting/ReportViewerFrame.aspx?AccountID=<%=svCustAcctId%>&reportfile=App_Data/repLearnerCompletion.frx">
              <!--[[-->Completion Status Report (Online)<!--]]--></a></td>
            <td style="text-align: left;">
              <a class="green" href="javascript:toggle('div_R1');">Description</a>
              <div style="text-align: left; margin-left: 20px" id="div_R1" class="div">
                <!--[[-->Lists completions/incompletions for all<br />programs assigned to your Learners.<!--]]-->
              </div>
            </td>
          </tr>


          <tr>
            <td>&nbsp; <a <%=fstatx%> target="_blank" href="/Gold/vuclientReporting/ReportExport.aspx?AccountID=<%=svCustAcctId%>&reportfile=repLearnerCompletionExport.frx&type=CSV">
              <!--[[-->Completion Status Report (CSV)<!--]]--></a></td>
            <td style="text-align: left;">
              <a class="green" href="javascript:toggle('div_R2');">Description</a>
              <div style="text-align: left; margin-left: 20px" id="div_R2" class="div">
                <!--[[-->Creates a CSV file of completions/incompletions for all<br />programs assigned to your Learners<!--]]-->
              </div>
            </td>
          </tr>

          <tr>
            <td>&nbsp; <a <%=fstatx%> target="_blank" href="/Gold/vuclientreporting/ReportExport.aspx?AccountID=<%=svCustAcctId%>&reportfile=repLearnerIncompleteCourseExport.frx&type=CSV"><!--[[-->InCompletion Report (CSV)<!--]]--></a></td>
            <td style="text-align: left;">
              <a class="green" href="javascript:toggle('div_R3');">Description</a>
              <div style="text-align: left; margin-left: 20px" id="div_R3" class="div">
                <!--[[-->Creates a CSV file of programs that have not been completed<br />for all programs assigned to your Learners<!--]]-->
              </div>
            </td>
          </tr>

          <tr>
            <td>&nbsp; <a <%=fstatx%> href="/Gold/vuClientReporting/AssReportFilter03.aspx?AccountID=<%=svCustAcctId%>">Assessment Response Report (Details)</a></td>
            <td>New&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="/Gold/vuClientReportingDev/AssReportFilter02.aspx?AccountID=<%=svCustAcctId%>">Assessment Response Report (Summary)</a></td>
            <td>New&nbsp;</td>
          </tr>

          <%   If svMembLevel = 5 Then %>
          <tr>
            <td style="vertical-align: top">&nbsp;&nbsp;<!--[[-->My Custom Policies<!--]]--></td>
            <td style="text-align: left;">
              <a class="green" href="javascript:toggle('div_R4');">Description</a>
              <div style="text-align: left; margin-left: 10px" id="div_R4" class="div">
                These are the policies that will be rendered from the &quot;smartLinks&quot; in your content from this Account.<br />
                &nbsp; <a href="#" onclick="docWindow('<%=fDocumentUrl("harassment.pdf", "", svLang, Left(svCustId, 4), svCustAcctId, "", "")%>')">harassment.pdf</a><br />
                &nbsp; <a href="#" onclick="docWindow('<%=fDocumentUrl("conflict.pdf", "", svLang, Left(svCustId, 4), svCustAcctId, "", "")%>')">conflict.pdf</a><br />
                &nbsp; <a href="#" onclick="docWindow('<%=fDocumentUrl("reaffirmation.pdf", "", svLang, Left(svCustId, 4), svCustAcctId, "", "")%>')">reaffirmation.pdf</a><br />
                &nbsp; <a href="#" onclick="docWindow('<%=fDocumentUrl("ethicsemployees.pdf", "", svLang, Left(svCustId, 4), svCustAcctId, "", "")%>')">ethicsemployees.pdf</a>
              </div>
            </td>
          </tr>
          <%   End If %>

          <% End If %>
        </table>
      </td>
    </tr>
    <% End If %>


    <!------------------------------------------------------- Manager -->
    <% If svMembLevel > 3 Then %>
    <tr id="t2" class="c1">
      <th class="underline" style="text-align: left">&nbsp;</th>
    </tr>
    <tr>
      <td colspan="2" class="c2">Advanced Services</td>
    </tr>
    <tr>
      <td colspan="2">
        <table class="table">
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="EcomReport.asp">Ecommerce Report Basic</a></td>
            <td>&nbsp;</td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="EcomReport0.asp">Ecommerce Report Advanced (Online)</a></td>
          </tr>



          <% If svMembLevel = 4 And Instr("CCHS ERGP IAPA VUBZ", Left(svCustId, 4)) > 0 Then 
               parms = "custId=" & svCustId & "&membNo=" & svMembNo & "&pageId=" & "ecommerceReport" & "&lang=" & Lcase(svLang)
               url = "/excel?profile=excel&parms=" & fBase64(parms)
          %>
          <tr>
            <td>&nbsp; <a target="_blank" <%=fstatx%> href="<%=url%>">Ecommerce Report Advanced (Excel)</a> for <%=Left(svCustId, 4)%></td>
            <td></td>
          </tr>
          <% End If  %>


          <% If svMembLevel = 5 Then 
               parms = "custId=VUBZ5678&membNo=" & fMembNo("5678", vPassword4) & "&pageId=" & "ecommerceReport" & "&lang=" & Lcase(svLang)
               url = "/excel?profile=excel&parms=" & fBase64(parms)
          %>
          <tr>
            <td>&nbsp; <a target="_blank" <%=fstatx%> href="<%=url%>">Ecommerce Report Advanced (Excel)</a> for VUBZ</td>
            <td></td>
          </tr>
          <% End If  %>

          <% If svMembLevel = 5 Then 
               parms = "custId=CCHS2544&membNo=" & fMembNo("2544", vPassword4) & "&pageId=" & "ecommerceReport" & "&lang=" & Lcase(svLang)
               url = "/excel?profile=excel&parms=" & fBase64(parms)
          %>
          <tr>
            <td>&nbsp; <a target="_blank" <%=fstatx%> href="<%=url%>">Ecommerce Report Advanced (Excel)</a> for CCHS</td>
            <td></td>
          </tr>
          <% End If  %>

          <% If svMembLevel = 5 Then 
               parms = "custId=ERGP2962&membNo=" & fMembNo("2962", vPassword4) & "&pageId=" & "ecommerceReport" & "&lang=" & Lcase(svLang)
               url = "/excel?profile=excel&parms=" & fBase64(parms)
          %>
          <tr>
            <td>&nbsp; <a target="_blank" <%=fstatx%> href="<%=url%>">Ecommerce Report Advanced (Excel)</a> for ERGP</td>
            <td></td>
          </tr>
          <% End If  %>

          <% If svMembLevel = 5 Then 
               parms = "custId=IAPA2859&membNo=" & fMembNo("2859", vPassword4) & "&pageId=" & "ecommerceReport" & "&lang=" & Lcase(svLang)
               url = "/excel?profile=excel&parms=" & fBase64(parms)
          %>
          <tr>
            <td>&nbsp; <a target="_blank" <%=fstatx%> href="<%=url%>">Ecommerce Report Advanced (Excel)</a> for IAPA</td>
            <td></td>
          </tr>
          <% End If  %>

          <% If svMembLevel = 5 Then 
               parms = "custId=CAAM3001&membNo=" & fMembNo("3001", vPassword4) & "&pageId=" & "ecommerceReport" & "&lang=" & Lcase(svLang)
               url = "/excel?profile=excel&parms=" & fBase64(parms)
          %>
          <tr>
            <td>&nbsp; <a target="_blank" <%=fstatx%> href="<%=url%>">Ecommerce Report Advanced (Excel)</a> for MPC (CAAM)</td>
            <td></td>
          </tr>
          <% End If  %>

          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="EcomReport1.asp">Ecommerce Sales Summary Report</a></td>
          </tr>
          <%    If vMemb_Ecom Then %>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="Ecom2Start.asp?vMode=More&vContentOptions=<%=vContentOptions%>">Post Manual Ecommerce Sales</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="EcomExtend.asp">Extend Access for Ecommerce Purchase</a></td>
          </tr>
          <%    End If %>
          <%    If svMembId = "VUV5_MGR" Or svMembLevel = 5 Then %>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="SeatThreshold.asp">Seat Threshold Report</a></td>
            <td></td>
          </tr>
          <%    End If %>
          <%    If fIsCorporate Then %>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="AccessReport.asp">Programs Assigned Report - Corporate Sites</a></td>
          </tr>
          <%    End If %>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="EcomReport3.asp">Program Sales Report</a></td>
          </tr>
          <%    If vCust_Id = "CAAM3001" Or svMembLevel = 5 Then %>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="EcomReport4.asp">Ecommerce Completion Report</a></td>
            <td>&nbsp;</td>
          </tr>
          <%    End If %>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="CatalogueDump.asp">Catalogue Dump</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="CatlByCustId_x.asp">Catalogue Nos (Excel)</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="CustomerExpires.asp">Maintain Customer Expiry Date</a></td>
          </tr>
          <% If fIsParent Then %>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="CustomerExpiryReport.asp">Channel Expiry Report</a></td>
            <td></td>
          </tr>
          <% End If %>

          <%    If svCustIssueIds Then %>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="IssueIds.asp">Generate Multiple Passwords</a></td>
          </tr>
          <%    End If %>

          <%    If svMembLevel = 5 And Not fNoValue(vCust_IssueIdsTemplate) Then %>
          <%    End If %>


          <%  If svMembLevel = 5 Or (svMembLevel = 4 And vMemb_MyWorld) Then %>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="TaskEdit1.asp">Maintain Tasks for My Learning</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="CritEdit.asp">Maintain Group 1 Table</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="JobsEdit.asp">Maintain Jobs Table</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="SkilEdit.asp">Maintain Skills Table</a></td>
          </tr>
          <%  End If %>



          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="UsersBulkInput.asp">Upload Learners Basic</a></td>
          </tr>
          <%  If svMembLevel > 3 Then %>

          <%    If (Instr("CMSS2592 UGRC1464", svCustId) > 0) Then %>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="<%=svDomain%>/Repository/Import/<%=svCustId%>.asp">Upload Learners Custom</a></td>
          </tr>
          <%    Else %>

          <%      
                If fIsV8 And vCust_ChannelReportsTo Then %>

          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="<%=svDomain%>/Repository/Upload3/Upload3.asp">Upload Learners Advanced with ReportsTo</a></td>
          </tr>

          <%      Else %>

          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="<%=svDomain%>/Repository/Upload2/Upload2.asp">Upload Learners Advanced</a></td>
          </tr>

          <%      End If  %>


          <%    End If %>
          <%  End If %>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="DocumentUpload.asp">Upload Custom Document</a> </td>
            <td>For content smartlinks</td>
          </tr>
        </table>
      </td>
    </tr>
    <% End If %>


    <!------------------------------------------------------- Customer Management System-->
    <% If svMembLevel = 5 Or (svMembLevel = 4 And vMemb_VuBuild) Then %>
    <tr id="t4" class="c1">
      <th class="underline" style="text-align: left">&nbsp;</th>
    </tr>
    <tr>
      <td class="c2">Customer Management</td>
    </tr>
    <tr>
      <td>
        <table class="table">
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="Catalogue.asp">Maintain Catalogue</a></td>
            <td>Updated</td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="Customers.asp?vEditCustId=<%=svCustId%>&vHidden=n">Maintain Customers</a></td>
            <td>Updated</td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="Customer.asp?vEditCustId=<%=svCustId%>&vHidden=n">Customer Profile</a></td>
            <td>Updated</td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="CatlFind.asp">Find Programs or Modules in Catalogues</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="CustomerActivityReport.asp">Customer Activity Report</a></td>
          </tr>
          <%   If svMembLevel = 5 Then %>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="CustomerSellerOwners.asp">Customer Sellers | Owners Report</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="LogsHistoryTransfer.asp">Transfer Learner History</a></td>
          </tr>
          <%   End If %>
        </table>
      </td>
    </tr>
    <% End If %>


    <!------------------------------------------------------- Content Management System -->
    <% If svMembLevel = 5 Or vMemb_LCMS Then %>
    <tr id="t5" class="c1">
      <th class="underline" style="text-align: left">&nbsp;</th>
    </tr>
    <tr>
      <td class="c2">Content Management</td>
    </tr>
    <tr>
      <td colspan="2">
        <table class="table">
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="Programs.asp">Maintain Programs</a></td>
            <td>Updated</td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="Modules.asp">Maintain Modules</a></td>
            <td>Updated</td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="ModSkillSet.asp">Update Module Skill Sets</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="javascript:jconfirm('ProgramUpdate.asp','OK to Continue?  This job can takes a long time to complete.')">Update Program Lengths</a></td>
          </tr>
        </table>
      </td>
    </tr>
    <% End If %>


    <!------------------------------------------------------- Vubiz Internal-->
    <% If svMembLevel = 5 Then %>
    <tr id="t6" class="c1">
      <th class="underline" style="text-align: left">&nbsp;</th>
    </tr>
    <tr>
      <td class="c2">Vubiz Internal</td>
    </tr>
    <tr>
      <td colspan="2">
        <table class="table">
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="Channel_Management.asp">Channel Analysis</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="CertificateSales.asp">Certificate Sales Report</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="/V5/Repository/Documents/MultiUserManual/UploadMultiUserManual.asp">Upload Multi User Manuals</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="/V5/RTEResetAttempts.htm">Reset Attempts on Scorm RTE (Temp)</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="/V5/Repository/V5_Vubz/0000/Tools/ERGP/Customers.asp">ERGP - Customer Control Center</a></td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="EcomHistory.asp">Ecommerce Transactions NOT Processed</a>&nbsp; </td>
            <td style="text-align: left;">Updated (work with Peter)</td>
          </tr>
          <% If Len(Trim(vCust_ParentId)) = 0 Then %>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="SnapshotExam.asp">Course Completion Snapshot</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="Patience.asp?vNext=SnapshotAll.asp">Course Usage Snapshot</a></td>
          </tr>
          <% End If %>
          <tr>
            <td colspan="2">&nbsp; <a href="EcomMemo.asp">Post Memo Funds to InternetSecure</a></td>
          </tr>
          <% If svMembInternal Then %>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="Patience.asp?vNext=ManageAdmins.asp">Manage <%=svCustId%> Administrators</a></td>
            <td>For Helen/Rosie</td>
          </tr>
          <tr>
            <td>&nbsp; <a <%=fstatx%> href="Patience.asp?vNext=ManageAuthors.asp">Manage <%=svCustId%> Authors</a></td>
            <td>For Rosie</td>
          </tr>
          <% End If %>

          <tr>
            <td>&nbsp; <a href="../vuNews.asp">vuNews Email Addresses</a></td>
            <td style="text-align: left;">Excel for Jen</td>
          </tr>
          <tr>
            <td>&nbsp; <a href="/V5/certLogs.aspx?svMembNo=<%=svMembNo%>">Certificate Logs</a></td>
            <td style="text-align: left;">For Lori</td>
          </tr>

          <tr>
            <td>&nbsp; <a href="ElavonRePost.asp">Elavon Repost</a></td>
            <td style="text-align: left;">For Lori when order is not completed</td>
          </tr>

          <tr>
            <td>&nbsp; <a target="_blank" href="/v6/memberChangeId.aspx">Change Learner/Facilitator Passwords</a></td>
            <td style="text-align: left;">Helps test ecom/ecom "P" transactions </td>
          </tr>

        </table>
      </td>
    </tr>
    <% End If %>


    <!------------------------------------------------------- Completion Testing -->
    <% If bCompletion Then %>
    <tr id="t8" class="c1">
      <th class="underline" style="text-align: left">&nbsp;</th>
    </tr>
    <tr>
      <td class="c2">Completion System</td>
    </tr>
    <tr>
      <td colspan="2">
        <table class="table">
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="Completion.asp">Completion Reports</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="Completion_Learners.asp">Learner Report</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="Completion_LocationManager.asp">Location Manager</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="Completion_LearnerLocation.asp?vMemb_No=<%=svMembNo%>">Learner Location</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="Completion_RoleManager.asp">Role Manager</a></td>
          </tr>
          <tr>
            <td colspan="2">&nbsp; <a <%=fstatx%> href="Completion_AddScores.asp">Add Scores</a></td>
          </tr>

          <tr>
            <td>&nbsp; <a <%=fstatx%> href="Completion_ReportSet.asp">Advanced Reports</a></td>
            <td style="text-align: left;">Bryan</td>
          </tr>
        </table>
      </td>
    </tr>
    <% End If %>
  </table>

  <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

</body>

</html>
