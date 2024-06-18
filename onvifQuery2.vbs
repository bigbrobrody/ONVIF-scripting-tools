' TODO get end points and namespaces for anything other than the device service from the GetServices reponse
' by mr rekcuf, modified by Ben Brody
' vbscript to query onvif camera
' Usage - one of the following:
' cscript /nologo onvifQuery2.vbs
' cscript /nologo onvifQuery2.vbs camera_IP userID password
'
Const ForReading = 1, ForWriting = 2, ForAppending = 8, CreateIfNeeded = true
dim fname 
set fso = CreateObject("Scripting.FileSystemObject")
last_ip = ""
last_user = ""
last_password = ""
outPath = "output"

if wscript.arguments.Count < 3 then
	' Get defaults
	If fso.FileExists("lastDevice.txt") then
		set file = fso.OpenTextFile("lastDevice.txt", ForReading)
		last_device = Split(file.ReadAll, vbcrlf)
		last_ip = last_device(0)
		last_user = last_device(1)
		last_password = last_device(2)
	End If

	' Get details from user
	ip = inputbox("Device IP address", "IP address", last_ip)
	user = inputbox("Device ONVIF username", "Username", last_user)
	password = inputbox("Device ONVIF password", "Password", last_password)

	' Save details for next time
	set file = fso.OpenTextFile("lastDevice.txt", ForWriting, CreateIfNeeded)
	file.write ip & vbcrlf & user & vbcrlf & password
	file.close

	' wscript.echo vbCrLf & "Usage: " & vbCrLf & " cscript /nologo " & wscript.scriptName  & " camera_IP userID password" & vbCrLf
	' wscript.quit
else
	ip = wscript.arguments.item(0)
	user =  wscript.arguments.item(1)
	password = wscript.arguments.item(2)
end if

' Function to pretty-print XML by adding whitespace between tags
Function prettyXml(ByVal sDirty)
    ' Put whitespace between tags (required for XSL transformation)
    sDirty = Replace(sDirty, "><", ">" & vbCrLf & "<")
    
    ' Create an XSL stylesheet for transformation
    Dim objXSL : Set objXSL = WScript.CreateObject("Msxml2.DOMDocument")
    objXSL.loadXML "<xsl:stylesheet version=""1.0"" xmlns:xsl=""http://www.w3.org/1999/XSL/Transform"">" & _
                    "<xsl:output method=""xml"" indent=""yes""/>" & _
                    "<xsl:template match=""/"">" & _
                    "<xsl:copy-of select="".""/>" & _
                    "</xsl:template>" & _
                    "</xsl:stylesheet>"
    
    ' Transform the XML
    Dim objXML : Set objXML = WScript.CreateObject("Msxml2.DOMDocument")
    objXML.loadXml sDirty
    objXML.transformNode objXSL
    
    prettyXml = objXML.xml
End Function

Function SOAPRequest(ByVal xml)
	xmlstd = _
	"xmlns:s='http://www.w3.org/2003/05/soap-envelope' " +_
	"xmlns:a='http://www.w3.org/2005/08/addressing'" 

	SOAPRequest = _
	"<?xml version='1.0' encoding='utf-8'?>" +_
	"<s:Envelope " + xmlstd + ">" +_
	"<s:Body " +_
	"xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " +_  
	"xmlns:xsd='http://www.w3.org/2001/XMLSchema'>" +_  
	xml +_
	"</s:Body>" +_
	"</s:Envelope>"

	SOAPRequest = Replace(SOAPRequest, "'", chr(34))
End Function

Function ONVIFExchange(service, message, profile)
	dim exchange
	set exchange = CreateObject("Scripting.Dictionary")

	url = "http://" & ip & "/onvif/" & service

	command = commands(message)
	if profile <> "" then
		command = replace(command, "REPLACEPROFILE", profile)
	end if

	Set xmlDoc = CreateObject("MSXML2.DOMDocument.6.0")
	xmlDoc.async = False
	xmlDoc.validateOnParse   = False
	xmlDoc.resolveExternals  = False

	namespaces = ""
	For Each namespace in ns
		namespaces = namespaces & namespace & " "
	Next
	xmlDoc.setProperty "SelectionNamespaces", namespaces

	with CreateObject("MSXML2.ServerXMLHTTP.6.0")
		.open "POST", url, False , user, password
		.setRequestHeader "Content-Type", "application/soap+xml; charset=utf-8"
		.setRequestHeader "Connection", "keep-alive"

		lResolve = 30 * 1000
		lConnect = 60 * 1000
		lSend = 30 * 1000
		lReceive = 120 * 1000
		.setTimeouts lResolve, lConnect, lSend, lReceive
		
		xml = SOAPRequest(command)

		WScript.Echo "Sending to " & url & vbCrLf & prettyXml(xml) & vbCrLf

		on error resume next
		.send xml
		xmlDoc.loadXML(.responseText)
		httpCode = "HTTP " & .Status & " " & .StatusText
		on error goto 0

	end with

	If Err.Number = 0 Then 
		WScript.Echo "Got:" & vbCrLf & httpCode & vbCrLf & prettyXml(xmlDoc.xml) & vbCrLf
	elseif  Err.Number = -2147012889 then
		wscript.Echo "Invalid IP Address or Hostname. Error code: " + hex(Err.Number)
	else
		wscript.Echo "Error code: " + hex(Err.Number)
	end if

	exchange.Add "url", url
	exchange.Add "message", message
	exchange.Add "service", service
	exchange.Add "request", xml
	exchange.Add "httpResponse", httpCode
	exchange.Add "response", xmlDoc
	exchange.Add "profile", profile

	set ONVIFExchange = exchange
End Function

Function writeToFiles(exchange, base_fname)
	dim fname
	fname = base_fname & "_" & exchange("service") & "_" & exchange("message")
	if exchange("profile") <> "" then
		fname = fname & "_" & exchange("profile")
	end if
	writeToFile fname & ".xml", prettyXml(exchange("request"))
	writeToFile fname & "_http_response.txt", "Response after posting to:" & vbCrLf & exchange("url") & vbCrLf & exchange("httpResponse")
	writeToFile fname & "Response.xml", prettyXml(exchange("response").xml)
End Function

Function writeToFile(outFile, string)
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	If Not objFSO.FolderExists(outPath) Then
		objFSO.CreateFolder(outpath)
	End If
	' Create the text file
	Set objFile = objFSO.CreateTextFile(objFSO.BuildPath(outPath, outFile), True)
	' Write a test string to the file
	objFile.Write string
	' Close the file
	objFile.Close
End Function

dim ns(6)
ns(0) = "xmlns:soap=""http://www.w3.org/2003/05/selope/"""
ns(1) = "xmlns=""http://schemas.microsoft.com/sharepoint/soap/"""
ns(2) = "xmlns:tt=""http://www.onvif.org/ver10/schema"""
ns(3) = "xmlns:trt=""http://www.onvif.org/ver10/media/wsdl"""
ns(4) = "xmlns:tr2=""http://www.onvif.org/ver20/media/wsdl"""
ns(5) = "xmlns:tds=""http://www.onvif.org/ver10/device/wsdl"""

Dim commands
Set commands = CreateObject("Scripting.Dictionary")

commands.Add "GetDeviceInformation", _
	"<GetDeviceInformation xmlns='http://www.onvif.org/ver10/device/wsdl'/>"

commands.Add "GetServices", _
	"<GetServices xmlns='http://www.onvif.org/ver10/device/wsdl'>" +_
		"<IncludeCapability>true</IncludeCapability>" +_
	"</GetServices>"

commands.Add "GetCapabilities", _
	"<GetCapabilities xmlns='http://www.onvif.org/ver10/device/wsdl'/>"

commands.Add "GetProfiles", _
	"<GetProfiles xmlns='http://www.onvif.org/ver10/media/wsdl'/>"

commands.Add "GetStreamUri", _
	"<GetStreamUri xmlns='http://www.onvif.org/ver20/media/wsdl'>" +_
		"<Protocol>RtspOverHttp</Protocol>" +_
		"<ProfileToken>REPLACEPROFILE</ProfileToken>" +_
	"</GetStreamUri>"

commands.Add "GetNodes", _
	"<GetNodes xmlns='http://www.onvif.org/ver20/ptz/wsdl'/>"

commands.Add "GetOptions", _
	"<GetOptions xmlns='http://www.onvif.org/ver20/imaging/wsdl'/>"

commands.Add "GetMoveOptions", _
	"<GetMoveOptions xmlns='http://www.onvif.org/ver20/imaging/wsdl'/>"

xxxx= _
"xmlns:SOAP-ENC='http://www.w3.org/2003/05/soap-encoding' " +_  
"xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance' " +_  
"xmlns:xsd='http://www.w3.org/2001/XMLSchema' " +_  
"xmlns:chan='http://schemas.microsoft.com/ws/2005/02/duplex' " +_  
"xmlns:c14n='http://www.w3.org/2001/10/xml-exc-c14n#' " +_
"xmlns:wsu='http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-utility-1.0.xsd' " +_
"xmlns:xenc='http://www.w3.org/2001/04/xmlenc#' " +_
"xmlns:wsc='http://schemas.xmlsoap.org/ws/2005/02/sc' " +_
"xmlns:ds='http://www.w3.org/2000/09/xmldsig#' " +_
"xmlns:wsse='http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd' " +_
"xmlns:xmime5='http://www.w3.org/2005/05/xmlmime' " +_
"xmlns:xmime='http://tempuri.org/xmime.xsd' " +_
"xmlns:xop='http://www.w3.org/2004/08/xop/include' " +_
"xmlns:wsrfbf='http://docs.oasis-open.org/wsrf/bf-2' " +_
"xmlns:wstop='http://docs.oasis-open.org/wsn/t-1' " +_
"xmlns:wsrfr='http://docs.oasis-open.org/wsrf/r-2' " +_
"xmlns:wsnt='http://docs.oasis-open.org/wsn/b-2' " +_
"xmlns:tt='http://www.onvif.org/ver10/schema' " +_
"xmlns:tds='http://www.onvif.org/ver10/device/wsdl' " +_
"xmlns:tev='http://www.onvif.org/ver10/events/wsdl' " +_
"xmlns:tptz='http://www.onvif.org/ver20/ptz/wsdl' " +_
"xmlns:trt='http://www.onvif.org/ver20/media/wsdl' " +_
"xmlns:timg='http://www.onvif.org/ver20/imaging/wsdl' " +_
"xmlns:tmd='http://www.onvif.org/ver10/deviceIO/wsdl' " +_
"xmlns:tns1='http://www.onvif.org/ver10/topics' " +_
"xmlns:ter='http://www.onvif.org/ver10/error' " +_
"xmlns:tds='http://www.onvif.org/ver10/device/wsdl' " +_
"xmlns:tnsaxis='http://www.axis.com/2009/event/topics' "


dim exchange

' Get device information and set output filename to manufacturer and model
set exchange = ONVIFExchange("device_service", "GetDeviceInformation", "")
manufacturer = exchange("response").selectNodes("//tds:Manufacturer")(0).text
model = exchange("response").selectNodes("//tds:Model")(0).text
fname = manufacturer & "_" & model
writeToFiles exchange, fname

set exchange = ONVIFExchange("device_service", "GetServices", "")
writeToFiles exchange, fname

set exchange = ONVIFExchange("device_service", "GetCapabilities", "")
writeToFiles exchange, fname

set exchange = ONVIFExchange("ptz_service", "GetNodes", "")
writeToFiles exchange, fname

' Get profiles and stream URI for each one
dim profile(10), streamUri(10)
index= 0
set exchange = ONVIFExchange("device_service", "GetProfiles", "")	' Should be media or media 2 service
writeToFiles exchange, fname
Set items = exchange("response").selectNodes("//trt:Profiles")
WScript.Echo "Found " & items.length & " Profile(s)." & vbCrLf
x = 0
y = 0 
For Each item In items
	profile(x) = item.getAttribute("token")
	WScript.Echo "Profile " & x & ": Token='" & item.getAttribute("token") & "' Name='" & item.selectNodes("tt:Name")(0).text & "'" & vbCrLf

	set exchange = ONVIFExchange("device_service", "GetStreamUri", profile(x))	' Should be media or media 2 service
	writeToFiles exchange, fname
	Set itemStreams = exchange("response").selectNodes("//tr2:Uri")
	For Each itemStream In itemStreams
		streamUri(y) = itemStream.text
			WScript.Echo "Found stream: " & itemStream.text & vbCrLf
		y = y + 1
	Next
	if streamUri(y) <> "" then exit for 
	x = x + 1
Next

' Get imaging options
set exchange = ONVIFExchange("imaging_service", "GetOptions", "")
writeToFiles exchange, fname

set exchange = ONVIFExchange("imaging_service", "GetMoveOptions", "")
writeToFiles exchange, fname

WScript.Quit
