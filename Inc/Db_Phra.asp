<%
  '____ Phra  ________________________________________________________________________

  Dim vPhra_No, vPhra_EN, vPhra_FR, vPhra_ES, vPhra_Pages, vPhra_Hidden 
  Dim vPhra_Eof, vPhraSql
  
  Dim p0, p1, p2, p3, p4, p5, p6, p7, p8, p9 '...special variants to embed values into 

  '...Define basic button names
  Dim bBack, bAdd, bContinue, bEdit, bEnable, bDisable, bGo, bNext, bReturn, bSubmit, bDelete, bUpdate, bLogOff, bFinished, bBegin, bDetails, bRestart, bPrint, bClose, bModify, bOnline
  Dim bNuggets, bTemplates, bInactivate, bSignin, bApply, bPreview, bSend, bDisplayAll, bHideAll, bClear, bSaveContinue, bSave, bReview, bRetry, bResend
  
  Select Case svLang
 
     Case "FR"
      bAdd          = "Ajouter"
      bApply        = "Appliquer"
      bBack         = "Précédente"
      bBegin        = "Commencer"
      bClear        = "Remettre à zéro"
      bClose        = "Fermer"
      bContinue     = "Continuer"
      bDelete       = "Supprimer"
      bDetails      = "Détails"
      bDisable      = "Désactiver"
      bDisplayAll   = "Afficher tout"
      bEdit         = "Éditer"
      bEnable       = "Activer"
      bFinished     = "Terminé"
      bGo           = "Aller"
      bHideAll      = "Masquer tout"
      bInactivate   = "Rendre inactif"
      bLogOff       = "Fermer la session"
      bModify       = "Modifier"
      bNext         = "Suivante"
      bNuggets      = "Capsules"
      bOnline       = "En ligne"
      bPreview      = "Aperçu"
      bPrint        = "Imprimer"
      bResend       = "Renvoyer"
      bRestart      = "Recommencer"
      bReview       = "Vérifier"
      bRetry        = "Refaire"
      bReturn       = "Retourner"
      bSave         = "Enregistrer"
      bSaveContinue = "Enregistrer et continuer"
      bSend         = "Envoyer"
      bSignIn       = "Ouvrir une session"
      bSubmit       = "Soumettre"
      bTemplates    = "Modèles"
      bUpdate       = "Mettre à jour"

    Case "ES"
      bAdd          = "Agregar"
      bApply        = "Aplicar"
      bBack         = "Anterior"
      bBegin        = "Comenzar"
      bClear        = "Eliminar"
      bClose        = "Cerrar"
      bContinue     = "Continuar"
      bDelete       = "Borrar"
      bDetails      = "Detalles"
      bDisable      = "Deshabilitar"
      bDisplayAll   = "Mostrar todo"
      bEdit         = "Corregir"
      bEnable			  = "Habilitar"
      bFinished     = "Terminado"
      bGo           = "Aller"
      bHideAll      = "Ocultar todo"
      bInactivate   = "Inactivar"
      bLogOff       = "Salga"
      bModify       = "Modificar"
      bNext         = "Siguiente"
      bNuggets      = "Pepitas"
      bOnline       = "En línea"
      bPreview      = "Vista previa"
      bPrint        = "Imprimir"
      bResend       = "Reenviar"
      bRestart      = "Reinicia"
      bRetry        = "Reintentar"
      bReturn       = "Retornar"
      bReview       = "Revisar"
      bSend         = "Enviar"
      bSave         = "Guardar"
      bSaveContinue = "Guardar y continuar"
      bSignIn       = "Conexión"
      bSubmit       = "Enviar"
      bTemplates    = "Plantillas"
      bUpdate       = "Actualizar"

    Case Else
      bAdd  			  = "Add"
      bApply        = "Apply"
      bBack         = "Back"
      bBegin        = "Begin"
      bClear        = "Clear"
      bClose        = "Close"
      bContinue     = "Continue"
      bDelete       = "Delete"
      bDetails      = "Details"
      bDisable      = "Disable"
      bDisplayAll   = "Display All"
      bEdit         = "Edit"
      bEnable			  = "Enable"
      bFinished     = "Finish"
      bGo           = "Go"
      bHideAll      = "Hide All"
      bInactivate   = "Inactivate"
      bLogOff       = "Log Off"
      bModify       = "Modifier"
      bNext         = "Next"
      bNuggets      = "Nuggets"
      bOnline       = "Online"
      bPreview      = "Preview"
      bPrint        = "Print"
      bResend       = "Resend"
      bRestart      = "Restart"
      bReview       = "Review"
      bRetry        = "Retry"
      bReturn       = "Return"
      bSave         = "Save"
      bSaveContinue = "Save and Continue"
      bSend         = "Send"
      bSignIn       = "Sign In"
      bSubmit       = "Submit"
      bTemplates    = "Templates"
      bUpdate       = "Update"
  End Select

 
 
  Sub sGetPhra (vPhraNo)
    vPhraSql = "SELECT * FROM Phra WHERE Phra_No =  " & vPhraNo
    sOpenDb4    
    Set oRs4 = oDb4.Execute(vPhraSql)
    If Not oRs4.Eof Then 
      sReadPhra
      vPhra_Eof = False
    Else
      vPhra_Eof = True
    End If
    Set oRs4 = Nothing
    sCloseDb4    
  End Sub



  Sub sGetHiddenPages (vPhraPage)
    vPhraSql = "SELECT Phra_No FROM Phra WHERE (CHARINDEX('" & vPhraPage & "', Phra_Hidden) > 0)"
    sOpenDb4
    Set oRs4 = oDb4.Execute(vPhraSql)
  End Sub


  Sub sReadPhra
    vPhra_No              = oRs4("Phra_No")
    vPhra_EN              = oRs4("Phra_EN")
    vPhra_FR              = fOkValue(oRs4("Phra_FR"))
    vPhra_ES              = fOkValue(oRs4("Phra_ES"))
    vPhra_Pages           = oRs4("Phra_Pages")
    vPhra_Hidden          = oRs4("Phra_Hidden")
  End Sub


  Sub sExtractPhra
    vPhra_No              = Request.Form("vPhra_No")
    vPhra_EN              = fUnquote(Request.Form("vPhra_EN"))
    vPhra_FR              = fUnquote(Request.Form("vPhra_FR"))
    vPhra_ES              = fUnquote(Request.Form("vPhra_ES"))
  End Sub
  

  Sub sUpdatePhra
    vPhraSql =        " UPDATE Phra SET"
    vPhraSql = vPhraSql & " Phra_EN       =  '" & vPhra_EN      & "'," 
    vPhraSql = vPhraSql & " Phra_FR       =  '" & vPhra_FR      & "',"
    vPhraSql = vPhraSql & " Phra_ES       =  '" & vPhra_ES      & "' "
    vPhraSql = vPhraSql & " WHERE Phra_No =   " & vPhra_No
'   sDebug
    sOpenDb4
    oDb4.Execute(vPhraSql)
    sCloseDb4
  End Sub


  '...this allows the site to call phrase that do not display hidden fields
  Function fPhra (vPhraNo)
    fPhra = fPhrase (vPhraNo, "Hide", "Visible")
  End Function


  '...this allows the site to call phrase that do not display hidden fields
  Function fPhraH (vPhraNo)
    fPhraH = fPhrase (vPhraNo, "Hide", "Hidden")
  End Function


  '...NEW - this allows the site to call a text phrase that comes from a DB
  Function fPhraId(vPhraId)
    vPhraSql = "SELECT * FROM Phra WHERE Phra_EN = '" & fUnQuote(vPhraId) & "'"
    sOpenDb4    
    Set oRs4 = oDb4.Execute(vPhraSql)
    If oRs4.Eof Then 
      fPhraId = vPhraId
    Else
      sReadPhra
      Select Case svLang
        Case "FR" : fPhraId = fOkValue(vPhra_FR)
        Case "ES" : fPhraId = fOkValue(vPhra_ES)
        Case Else : fPhraId = vPhra_EN
      End Select
    End If
    If Len(fPhraId) = 0 Then fPhraId = vPhraId
    Set oRs4 = Nothing
    sCloseDb4     
  End Function



  '...this allows you to decide if you want hidden or not (for footer.asp)
  Function fPhrase (vPhraNo, vShowHide, vVisible)

    Dim vLang

    '...to override session language add "_XX" at the end of the Phrase No, ie 27_FR
    If Len(vPhraNo) > 3 Then 
      If Instr("_EN _FR _ES ", Right(vPhraNo, 3)) > 0 Then
        vLang   = Right(vPhraNo, 2)
        vPhraNo = Left(vPhraNo, Len(vPhraNo)-2)
      End If
    End If

    '...otherwise use the session language
    If vLang = "" Then vLang = svLang
    If vLang = "" Then vLang = "EN"

    sGetPhra vPhraNo
    If vPhra_Eof Then
      fPhrase = ""
    Else
      Select Case svLang
        Case "FR" : fPhrase = fOkValue(vPhra_FR)
        Case "ES" : fPhrase = fOkValue(vPhra_ES)
        Case Else : fPhrase = vPhra_EN
      End Select
      '...if no foreign phrase on file then use English
      If Len(fPhrase) = 0 Then
        fPhrase = vPhra_EN & " *"
      End If

      '...embed special values into the thingme
      fPhrase = Replace(fPhrase, "^1", p1)
      fPhrase = Replace(fPhrase, "^2", p2)
      fPhrase = Replace(fPhrase, "^3", p3)
      fPhrase = Replace(fPhrase, "^4", p4)
      fPhrase = Replace(fPhrase, "^5", p5)
      fPhrase = Replace(fPhrase, "^6", p6)
      fPhrase = Replace(fPhrase, "^7", p7)
      fPhrase = Replace(fPhrase, "^8", p8)
      fPhrase = Replace(fPhrase, "^9", p9)
      fPhrase = Replace(fPhrase, "^0", p0)

      If Left(fPhrase, 1)  = vbCr Then fPhrase = Mid(fPhrase, 2)
      If Right(fPhrase, 1) = vbCr Then fPhrase = Left(fPhrase, Len(fPhrase)-1)
 

    End If  

    '...allow admins to edit phrases online if translate session is true
    If svTranslate and vVisible = "Visible" Then
      If vShowHide = "Show" Then
        fPhrase = "<a style='TEXT-DECORATION: none; font-size: 8pt; color: #FFA500' target='_self' href='/V5/Translate.asp?vPhraNo=" & vPhraNo & "&vNext=" & svPage & "'>o</a> " & Server.HtmlEncode(fPhrase)
      ElseIf vShowHide = "Hide" Then
        fPhrase = "<a style='TEXT-DECORATION: none; font-size: 8pt; color: #FFA500' target='_self' href='/V5/Translate.asp?vPhraNo=" & vPhraNo & "&vNext=" & svPage & "'>o</a> " & fPhrase
      End If        
    End If

  End Function  


  '...diplay constants in a table that can allow online editing - called from "Inc/Footer.asp"
  Sub sHiddenPhrases
    Dim vPhraNo
    vPhraSql = "SELECT Phra_No FROM Phra WHERE (CHARINDEX('" & svPage & "', Phra_Hidden) > 0)"
    sOpenDb5
    Set oRs5 = oDb5.Execute(vPhraSql)

    If oRs5.Eof Then
      Set oRs5 = Nothing
      sCloseDb5  
      Exit Sub
    End If
%>

    <div align="center">
      <table border="1" width="65%" id="table1" cellspacing="0" cellpadding="5" bgcolor="#FFFFFF" style="border-collapse: collapse" bordercolor="orange">
        <tr>
          <td>
          <table border="0" cellpadding="2" style="border-collapse: collapse" bgcolor="#FFFFFF" width="100%" id="table2">
            <tr>
              <td align="center" class="navTableHeader" height="30"><font color="#FFA500">Hidden Phrases</font> </td>
            </tr>
            <tr>
              <td>&nbsp;</td>
            </tr>
            <%
              Do While Not oRs5.Eof
                vPhraNo = oRs5("Phra_No")
            %>
            <tr>
              <td><font color="#FFA500"><%=fPhrase(vPhraNo, "Show", "Visible")%></font></td>
            </tr>
            <%
                oRs5.MoveNext	        
              Loop
              Set oRs5 = Nothing
              sCloseDb5             
            %>
            <tr>
              <td>&nbsp;</td>
            </tr>
          </table>
          </td>
        </tr>
      </table>
    </div>
    
<%
    End Sub
%>