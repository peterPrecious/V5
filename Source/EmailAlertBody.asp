<% 
  Sub sEmailAlert (vEmail_Notify, vSubject, vEmail_Note, vEmail_Type)
    
    vErr = ""

    Dim vFirstName, vLastName, vName, vEmail, vId, vBody, vSender, vRecipients
    Dim aEmail_Notify, iMax
    
    '_________________________________________________________________________
    '...put member nos in an array
    aEmail_Notify = Split(vEmail_Notify, ", ")
    iMax = Ubound(aEmail_Notify)
    For i = 0 to iMax
      vMemb_No = Clng(Trim(aEmail_Notify(i)))
      sGetMemb vMemb_No
      If NOT vMemb_Eof Then 
        ReDim Preserve aSendTo(4, i)
        aSendTo (0, i) = vMemb_Email
        aSendTo (1, i) = vMemb_FirstName
        aSendTo (2, i) = vMemb_LastName
        aSendTo (3, i) = 0
        aSendTo (4, i) = vMemb_Id
      End If
    Next

    '_________________________________________________________________________
    '...send out an array of names with a trailing note

    For i = 0 to iMax
      '...get fields table
      vFirstName         = aSendTo(1, i)
      vEmail             = aSendTo(0, i)
      vName              = aSendTo(1, i) & " " & aSendTo(2,i)
      vId                = aSendTo(4, i)

      '...build up message body and include the task title
      vBody              = "<br>"
      vBody              = vBody & "<!--{{-->Hello<!--}}-->" & " " & vFirstName & ",<br><br>"
      vBody              = vBody & vEmail_Type & " " & "<!--{{-->for you in task<!--}}-->" & ":<br>"
      vBody              = vBody & vEmail_TaskTitle & "<br><br>"

      '...generate a goto url (note svLang is added to override customer language field, when desired)
      vBody              = vBody & "//" & svHost & "/Goto.asp?vCode=" & fCreateUrls ("//" & svHost & "/default.asp?vCust=" & svCustId & "&vId=" & vId & "&vTskH_Id=" & vTskH_Id & "&vAction=MYWORLD" & "&vLang=" & svLang) & vbCrLf & vbCrLf 

      If vEmail_Note <> "" Then
        vBody            = vBody & "<br><br>" & vEmail_Note & "<br><br>"
      End If
      vBody              = vBody & "<br><br>" & "<!--{{-->Thank you<!--}}-->" & "<br>" & svMembFirstName

      '...feed parms to mail object (also uses session variables: svMembFirstName, svMembLastName, svMembEmail, svCustID)
      vSender            = svMembFirstName & " " & svMembLastName & " <" & svMembEmail & ">"
      vRecipients        = vName & " <" & vEmail & ">"
    
        '...If sent successfully then set flag
      vErr = fFathMail(vSubject, vBody, vSender, vRecipients) 
      If vErr = "Ok" Then aSendTo(3, i) = True
    Next

    '...check flag if any errors in array
    vErr = ""
    For i = 0 To iMax
      If Not aSendTo(3, i) Then 
        vErr = vErr & "<br>" & aSendTo (1, i) & " " & aSendTo(2, i)
      End If      
    Next

    If vErr <> "" Then 
      Server.Execute vShellHi 
%>
    <table border="1" width="100%" style="border-collapse: collapse" bordercolor="#DDEEF9" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center">
        <h6><br>
        <!--[[-->The following could not be alerted:<!--]]--> <br><%=vErr%> </h6>
        <h6>
        <!--[[-->Their email addresses may be invalid.<!--]]--> <br>
        <!--[[-->Please contact your facilitator.<!--]]--></h6>
        <p><a href="MyWorld.asp?vTskH_Id=<%=vTskH_Id%>"><img border="0" src="../Images/Icons/World.gif" width="34" height="24"></a><br>&nbsp;</p></td>
      </tr>
    </table>

    <!--#include virtual = "V5/Inc/Shell_Lo.asp"-->

<%
    End If
  End Sub
%>