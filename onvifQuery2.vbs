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

' Function ONVIFExchange(service, message, token)
Function ONVIFExchange(services, service_name, message, options, token)
	dim exchange
	set exchange = CreateObject("Scripting.Dictionary")

	' If media2 is not available fallback to media
	if service_name = "media2" and not services.exists(service_name) then service_name = "media"

	if not services.exists(service_name) then
		exchange.Add "url", "Error: ONVIF service not found"
		exchange.Add "message", message
		exchange.Add "service", service
		exchange.Add "request", "Error: ONVIF service not found"
		exchange.Add "httpResponse", "Error: ONVIF service not found"
		exchange.Add "response", "Error: ONVIF service not found"
		exchange.Add "token", token
		set ONVIFExchange = exchange
		return
	end if

	set service = services(service_name)

	url = service("XAddr")
	service_ns = service("ns")

	command = "<" & message & " xmlns=""" & service_ns & """>" & options & "</" & message & ">"
	if token <> "" then
		command = replace(command, "REPLACETOKEN", token)
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
		wscript.Echo "Invalid IP Address or Hostname. Error code: " + hex(Err.Number) & vbCrLf
	else
		wscript.Echo "Error code: " + hex(Err.Number) & vbCrLf
	end if

	exchange.Add "url", url
	exchange.Add "message", message
	exchange.Add "service", service_name
	exchange.Add "request", xml
	exchange.Add "httpResponse", httpCode
	exchange.Add "response", xmlDoc
	exchange.Add "token", token

	set ONVIFExchange = exchange
End Function

Function writeToFiles(exchange, base_fname)
	dim fname
	dim token: token = ""
	if exchange("token") <> "" then token = "_" & exchange("token")
	fname = base_fname & "_" & exchange("service") & token & "_" & exchange("message")
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

' Other namespaces to include above as required
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

' Initialise ONVIF services with device service
Dim services, service
Set services = CreateObject("Scripting.Dictionary")
Set service = CreateObject("Scripting.Dictionary")
service.Add "ns", "http://www.onvif.org/ver10/device/wsdl"
service.Add "XAddr", "http://" & ip & "/onvif/device_service"
services.Add "device", service

' Get device information and set output filename to manufacturer and model
set exchange = ONVIFExchange(services, "device", "GetDeviceInformation", "", "")
manufacturer = exchange("response").selectSingleNode("//tds:Manufacturer").text
manufacturer = Left(Replace(manufacturer, " ", "_"), 15)
model = exchange("response").selectSingleNode("//tds:Model").text
model = Left(Replace(model, " ", "_"), 15)
fname = manufacturer & "_" & model
writeToFiles exchange, fname

' Get services
opts = "<IncludeCapability>true</IncludeCapability>"
set exchange = ONVIFExchange(services, "device", "GetServices", opts, "")
writeToFiles exchange, fname

' Save details of all the services
Dim objRegEx
Set objRegEx = CreateObject("VBScript.RegExp")
With objRegEx
      .Pattern = "^\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,4}\b$" 
      .IgnoreCase = True 
End With
set xml_services = exchange("response").selectNodes("//tds:Service")
dim service_ns, service_version, service_name, service_XAddr
For each xml_service in xml_services
	Do
		Set service = CreateObject("Scripting.Dictionary")
		service_ns = xml_service.selectSingleNode("tds:Namespace").text
		objRegEx.Pattern = "ver([0-9]{2})\/"
		set matches = objRegEx.Execute(service_ns)
		if matches.count = 0 then exit do
		service_version = matches(0).SubMatches(0)
		objRegEx.Pattern = "[0-9]{2}\/([^\/]*)\/"
		set matches = objRegEx.Execute(service_ns)
		if matches.count = 0 then exit do
		service_name = matches(0).SubMatches(0)
		if service_name = "media" and service_version = "20" then
			service_name = "media2"
		end if
		service_XAddr = xml_service.selectSingleNode("tds:XAddr").text
		WScript.Echo "Service: name='" & service_name & "' version='" & service_version & "' namespace='" & service_ns & "' XAddr='" & service_XAddr & "'" & vbCrLf
		if service_name = "device" then exit do
		service.Add "ns", service_ns
		service.Add "XAddr", service_XAddr
		services.Add service_name, service
	Loop While False
Next

' Get capabilities
set exchange = ONVIFExchange(services, "device", "GetCapabilities", "", "")
writeToFiles exchange, fname

' Get PTZ nodes
set exchange = ONVIFExchange(services, "ptz", "GetNodes", "", "")
writeToFiles exchange, fname

' Get profiles
index= 0
Set exchange = ONVIFExchange(services, "media", "GetProfiles", "", "")
writeToFiles exchange, fname
Set items = exchange("response").selectNodes("//trt:Profiles")
count = 0: For Each item in items: count = count + 1: Next
WScript.Echo "Found " & count & " Profile(s)." & vbCrLf

' Get stream URI for each profile using media
opts =_
	"<StreamSetup>" +_
		"<Stream>RTP-Multicast</Stream>" +_
		"<Transport>" +_
			"<Protocol>RTSP</Protocol>" +_
		"</Transport>" +_
	"</StreamSetup>" +_
	"<ProfileToken>REPLACETOKEN</ProfileToken>"
For Each item In items
	token = item.getAttribute("token")
	WScript.Echo "Profile: Token='" & token & "' Name='" & item.selectSingleNode("tt:Name").text & "'" & vbCrLf

	set exchange = ONVIFExchange(services, "media", "GetStreamUri", opts, token)
	writeToFiles exchange, fname
Next

' Get stream URI for each profile using media2
If services.exists("media2") Then
	opts =_
		"<Protocol>RtspMulticast</Protocol>" +_
		"<ProfileToken>REPLACETOKEN</ProfileToken>"
	For Each item In items
		token = item.getAttribute("token")
		WScript.Echo "Profile: Token='" & token & "' Name='" & item.selectSingleNode("tt:Name").text & "'" & vbCrLf

		set exchange = ONVIFExchange(services, "media2", "GetStreamUri", opts, token)
		writeToFiles exchange, fname

		' streamUri = exchange("response").selectSingleNode("//tr2:Uri").text	' trt for media, tr2 for media2
		' WScript.Echo "Found stream: " & streamUri & vbCrLf
	Next
End If

' Get video sources
set exchange = ONVIFExchange(services, "media", "GetVideoSources", "", "")
writeToFiles exchange, fname
Set items = exchange("response").selectNodes("//trt:VideoSources")
count = 0: For Each item in items: count = count + 1: Next
WScript.Echo "Found " & count & " VideoSource(s)." & vbCrLf

' Get options for each video source
opts = "<VideoSourceToken>REPLACETOKEN</VideoSourceToken>"
For Each item in items
	token = item.getAttribute("token")
	WScript.Echo "Video source: Token='" & token & "'" & vbCrLf

	set exchange = ONVIFExchange(services, "imaging", "GetOptions", opts, token)
	writeToFiles exchange, fname

	set exchange = ONVIFExchange(services, "imaging", "GetMoveOptions", opts, token)
	writeToFiles exchange, fname
Next

WScript.Quit
