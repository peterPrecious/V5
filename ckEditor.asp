<%
  If (Request.Form.Count > 0 ) Then
    a = Request("progDesc")
    response.write Server.HTMLEncode(a) & "<br><br><br><br>"
  End If
%>





<!DOCTYPE html>
<!--
Copyright (c) 2003-2016, CKSource - Frederico Knabben. All rights reserved.
For licensing, see LICENSE.md or //ckeditor.com/license
-->
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
  <title>CKEditor Sample</title>
  <script src="/V5/Inc/jQuery.js"></script>
  <script src="ckEditor/ckEditor.js"></script>
  <script src="ckEditor/ckEditorVu.js"></script>
  <script>
    $(function(){initCkEditorVu();});
  </script>

</head>

<body>

  <form id="fProgram" method="POST" action="ckEditor.asp">
    <table style="width: 800px; margin: auto;">
      <tr>
        <th style="vertical-align: top;">Description : </th>
        <td>
          <textarea name="progDesc" id="editor" maxlength="8000">
            <p>The Fair Debt Collection Practices Act (FDCPA) is a Federal law that limits the behavior and actions of debt collectors who are attempting to collect the debt for another person or entity. The law restricts the means and methods by which the debtor can be contacted including the time of day the contact can be made.</p>
            <p>This course will provide you with an in-depth look at the Fair Debt Collection Practices Act. You will learn about the legislation that has been enacted to provide guidelines to you, as a debt collector, and protect the consumer.</p>
            <p>This course includes an examination.</p>
            <p><strong><em> </em></strong></p>
            <p><strong><em>Learning Objectives:</em></strong></p>
            <ul>
            <li>Explain the Fair Debt Collection Practices Act</li>
            <li>Understand the guidelines of fair debt collection and the legislation in place to protect the consumer</li>
            </ul>
            <ul>
            <li>Comply with the rules and regulations surrounding the Fair Debt Collection Practices Act</li>
            </ul>
            <p><strong><em>Course Outline:</em></strong></p>
            <ul>
            <li>SFindings, Purpose and Key Definitions</li>
            <li>     Exemption for State Regulation</li>
            </ul>
            <ul>
            <ul>
            <li>Background and Findings</li>
            <li>Definitions and Coverage</li>
            </ul>
            <li>Communication</li>
            <ul>
            <li>Acquisition of Location Information</li>
            <li>Communicating with Third Parties</li>
            <li>Communication Permitted with Consumers</li>
            <li>When to Cease Communication with Consumers</li>
            </ul>
            <li>Harassment, Abuse and Conduct</li>
            <ul>
            <li>Abuse and Harassment</li>
            <li>False or misleading representations</li>
            <li>Unfair Practices</li>
            </ul>
            <li>Validation of Debts, Multiple Debts</li>
            <ul>
            <li>Validation of Debts</li>
            <li>Multiple Debts</li>
            </ul>
            <li>Legal Actions, Civic Liability, Reports</li>
            <ul>
            <li>Legal Actions by Debt Collectors</li>
            <li>Furnishing Certain Deceptive Forms</li>
            <li>Civic Liability</li>
            <li>Defenses</li>
            <li>Jurisdiction and Statute of Limitations</li>
            <li>Reports to Congress by the Commission</li>
            </ul>
            <li>Enforcement, State Law, Exemption</li>
            <ul>
            <li>Administrative Enforcement</li>
            <li>Relation to State Law</li>
            </ul>
            </ul>
            <p><strong><em>Duration:</em></strong></p>
            <p>0.83 hours</p><br>
            <p><strong><em>Features:</em></strong></p>
            <p>Audio, hybrid</p><br>
            <p><strong><em>Module(s):</em></strong></p>
            <p>7976EN</p>
          </textarea>
        </td>
      </tr>
      <tr>
        <td colspan="2" style="text-align: center; padding-top: 30px;">
          <input type="submit" value="Update" /></td>
      </tr>
    </table>
  </form>
</body>
</html>
