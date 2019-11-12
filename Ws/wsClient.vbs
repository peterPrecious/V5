Dim oXml, oUrl, vContent
Dim vCust, Vid, vAction

'...instantiate XmlHttp object
Set oXml = CreateObject("Microsoft.XmlHttp")
Set oUrl = CreateObject("Ixsso.Util")

'...define parms to be passed (which will be our XML encoded file)
vCust    = "Vubz2307"
vId      = "Eggleston"
vAction  = "GetCatalogue" 

'...xml format
vContent = "<VUBIZ><WS vAction='" & vAction & "' vCust='" & vCust & "' vId='" & vId & "'/></VUBIZ>"

'...test sending nothin
'vContent = ""

'...open the channel to the WebService
oXml.Open "POST", "http://s2.vubiz.com:8000/V5/WSxml.asp", False

'...set Content Type and Length information
oXml.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
oXml.SetRequestHeader "Content-Length", Len(vContent)

'...Send form parameters to the Web Service
oXml.Send vContent

'...display the Response from the WebService (optional)
MsgBox oXml.ResponseText
