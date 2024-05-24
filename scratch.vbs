' GetDeviceInformation=_
' "<GetDeviceInformation xmlns='http://www.onvif.org/ver10/device/wsdl'/>"

' GetServices=_
' "<GetServices xmlns='http://www.onvif.org/ver10/device/wsdl'>" +_
' "<IncludeCapability>true</IncludeCapability>" +_
' "</GetServices>"

' GetProfiles=_
' "<GetProfiles xmlns='http://www.onvif.org/ver10/media/wsdl'/>"

' GetStreamUri=_
' "<GetStreamUri xmlns='http://www.onvif.org/ver20/media/wsdl'>" +_
' "<Protocol>RtspOverHttp</Protocol>" +_
' "<ProfileToken>REPLACEPROFILE</ProfileToken>" +_
' "</GetStreamUri>"

' GetNodes=_
' "<GetNodes xmlns='http://www.onvif.org/ver10/ptz/wsdl'>" +_
' "</GetNodes>"


' dim profile(10), streamUri(10)


' index= 0


' url = "http://" & ip & "/onvif/device_service"
' user =  wscript.arguments.item(1)
' password = wscript.arguments.item(2)

' with CreateObject("MSXML2.ServerXMLHTTP.6.0")

' 	.open "POST", url, False , user, password
' 	.setRequestHeader "Content-Type", "application/soap+xml; charset=utf-8"
' 	' .setRequestHeader "Accept-Encoding", "gzip, deflate"
' 	.setRequestHeader "Connection", "keep-alive"

' 	lResolve = 30 * 1000
' 	lConnect = 60 * 1000
' 	lSend = 30 * 1000
' 	lReceive = 120 * 1000
' 	.setTimeouts lResolve, lConnect, lSend, lReceive
	
' 	xml = SOAPRequest(GetDeviceInformation)
' 	WScript.Echo "Sending:" & vbCrLf & prettyXml(xml) & vbCrLf
' 	.send xml
' 	xmlDoc.loadXML(.responseText)
' 	httpCode = "HTTP " & .Status & " " & .StatusText
' 	WScript.Echo "Got:" & vbCrLf & httpCode & vbCrLf & prettyXml(xmlDoc.xml) & vbCrLf

' 	manufacturer = xmlDoc.selectNodes("//tds:Manufacturer")(0).text
' 	model = xmlDoc.selectNodes("//tds:Model")(0).text
' 	fname = manufacturer & "_" & model
' 	writeToFile fname & "_GetDeviceInformation.xml", prettyXml(xml)
' 	writeToFile fname & "_GetDeviceInformationResponse.xml", prettyXml(xmlDoc.xml)

' 	xml = SOAPRequest(GetProfiles)
' 	WScript.Echo "Sending:"
' 	WScript.Echo prettyXml(xml)
' 	WScript.Echo

' 	On Error Resume Next
' 	.open "POST", url, False , user, password
' 	.send xml

' 	xmlDoc.loadXML(.responseText)
' 	WScript.Echo "Got:"
' 	httpCode = "HTTP " & .Status & " " & .StatusText
' 	WScript.Echo httpCode
' 	WScript.Echo prettyXml(xmlDoc.xml)
' 	WScript.Echo


' 	If Err.Number = 0 Then 

' 		Set items = xmlDoc.selectNodes("//trt:Profiles")

' 		WScript.Echo "Found " & items.length & " Profile(s)."
' 		WScript.Echo
  	
' 		x = 0
' 		y = 0 
'   		For Each item In items
' 			profile(x) = item.getAttribute("token")
' 	    		WScript.Echo " Profile " &x &" token: " & item.getAttribute("token") &" Name: " & item.selectNodes("tt:Name")(0).text

' 			.open "POST", url, False , user, password
' 			.setRequestHeader "Content-Type", "application/soap+xml; charset=utf-8" 
' 			.setRequestHeader "Accept-Encoding", "gzip, deflate"
' 			.setRequestHeader "Connection", "keep-alive"
' 			xml = Replace(SOAPRequest(GetStreamUri), "REPLACEPROFILE", profile(x))
' 			WScript.Echo "Sending:"
' 			WScript.Echo prettyXml(xml)
' 			WScript.Echo
' 			.send xml

' 		   	xmlDoc.loadXML(.responseText)
' 			WScript.Echo "Got:"
' 			httpCode = "HTTP " & .Status & " " & .StatusText
' 			WScript.Echo httpCode
' 			WScript.Echo prettyXml(xmlDoc.xml)
' 			WScript.Echo
' 			Set itemStreams = xmlDoc.selectNodes("//tr2:Uri")
' 	  		For Each itemStream In itemStreams
' 				streamUri(y) = itemStream.text
' 	    			WScript.Echo "   Found stream: " &itemStream.text 
' 				y =y + 1
' 	  		Next
' 			if streamUri(y) <> "" then exit for 
' 			x = x +1
'   		Next

' 	elseif  Err.Number = -2147012889 then

' 		wscript.Echo "Invalid IP Address or Hostname. Error code: " +hex(Err.Number)

' 	else

' 		wscript.Echo "Error code: " + hex(Err.Number)

' 	end if

' 	On Error GoTo 0

' end with

dim completed

msgbox inputboxml("Enter text:", "Multiline inputbox via HTA", "Enter your name : " & vbcrlf & "Enter your Age : " & vbcrlf &_
"Enter email id you want to send : " & vbcrlf & "Enter Subject : " & vbcrlf & "Enter Email Body : " )

function inputboxml(prompt, title, defval)
    set window = createwindow()
    completed = 0
    defval = replace(replace(replace(defval, "&", "&amp;"), "<", "&lt;"), ">", "&gt;")
    with window
        with .document
            .title = title
            .body.style.background = "buttonface"
            .body.style.fontfamily = "consolas, courier new"
            .body.style.fontsize = "8pt"
            .body.innerhtml = "<div><center><nobr>" & prompt & "</nobr><br><br></center><textarea id='hta_textarea' style='font-family: consolas, courier new; width: 100%; height: 400px;'>" & defval & "</textarea><br><button id='hta_cancel' style='font-family: consolas, courier new; width: 85px; margin: 10px; padding: 3px; float: right;'>Cancel</button><button id='hta_ok' style='font-family: consolas, courier new; width: 85px; margin: 10px; padding: 3px; float: right;'>OK</button></div>"
        end with
        .resizeto 550, 550
        .moveto 100, 100
    end with
    window.hta_textarea.focus
    set window.hta_cancel.onclick = getref("hta_cancel")
    set window.hta_ok.onclick = getref("hta_ok")
    set window.document.body.onunload = getref("hta_onunload")
    do until completed > 0
        wscript.sleep 10
    loop
    select case completed
    case 1
        inputboxml = ""
    case 2
        inputboxml = ""
        window.close
    case 3
        inputboxml = window.hta_textarea.value
        window.close
    end select
end function

function createwindow()
    ' rem source http://forum.script-coding.com/viewtopic.php?pid=75356#p75356
    dim signature, shellwnd, proc
    on error resume next
    signature = left(createobject("Scriptlet.TypeLib").guid, 38)
    do
        set proc = createobject("WScript.Shell").exec("mshta ""about:<head><script>moveTo(-32000,-32000);</script><hta:application id=app border=dialog minimizebutton=no maximizebutton=no scroll=no showintaskbar=yes contextmenu=no selection=yes innerborder=no icon=""%windir%\system32\notepad.exe""/><object id='shellwindow' classid='clsid:8856F961-340A-11D0-A96B-00C04FD705A2'><param name=RegisterAsBrowser value=1></object><script>shellwindow.putproperty('" & signature & "',document.parentWindow);</script></head>""")
        do
            if proc.status > 0 then exit do
            for each shellwnd in createobject("Shell.Application").windows
                set createwindow = shellwnd.getproperty(signature)
                if err.number = 0 then exit function
                err.clear
            next
        loop
    loop
end function

sub hta_onunload
    completed = 1
end sub

sub hta_cancel
    completed = 2
end sub

sub hta_ok
    completed = 3
end sub

WScript.Quit

