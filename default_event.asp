<%
    Response.Expires = -1000
' ****************************************************************************
' File:         Default.asp
' Description:  Invoked by MIU-1000 XML History Log update.
' Need to use a local file "sample.xml" to test from a browser.
' 
' ****************************************************************************
' ****************************************************************************
' Sub:          GetUtcOffsetMinutes
' Description:  Gets the number of minutes between local time and UTC.
'
' Params:       None
' ****************************************************************************

Function GetUtcOffsetMinutes()
    Dim key
    key = "UtcOffsetMinutes"
    GetUtcOffsetMinutes = Application(key)
    If IsEmpty(GetUtcOffsetMinutes) Then
        'Create Shell object to read registry
        Dim oShell, atb, offsetMinutes
        Set oShell = CreateObject("WScript.Shell")
        'Reading the registry
        GetUtcOffsetMinutes = oShell.RegRead("HKEY_LOCAL_MACHINE\System\CurrentControlSet\Control\TimeZoneInformation\ActiveTimeBias")
        Application(key) = GetUtcOffsetMinutes
        Set oShell = Nothing
        
    End If
        
End Function
' ****************************************************************************
' Sub:          ConvertToLocalTime
' Description:  Converts the UTC time to local time.
'
' Params:       utcDateTime: The UTC time to convert to local time.
' ****************************************************************************
Function ConvertToLocalTime(utcDateTime)
    ConvertToLocalTime = DATEADD("n", -(GetUtcOffsetMinutes()), utcDateTime)
End Function


set objXML = Server.CreateObject("Microsoft.XMLDOM")
objXML.ValidateOnParse = True
objXML.async = False
' Request for real *script launched from MIU-1000* , Canned local file sample.xml for debug
objXML.Load(Request)
'objXML.Load("c:\inetpub\wwwroot\miu1000\sample.xml")

   If objXML.ParseError.errorCode <> 0 Then
		'response.write("<p>")	
		Response.Write("<ErrorList><Error>" & objXML.parseError.reason & "At line: " & objXML.parseError.line & "</Error></ErrorList>")
		'response.write("</p>")
		conn.close
	 Else
		Set objRootElement = objXML.documentElement
   End If

	dRecID = objXML.documentElement.getAttribute("ID")
	dRecName = objXML.documentElement.getAttribute("Name")
	dRecNoSensor = objXML.documentElement.getAttribute("NoSensors")

	Set objNode = objXML.documentElement.firstChild

   For Each xmlPNode In objRootElement.childNodes

		For Each xmlNode In xmlPNode.childNodes
			Select Case xmlNode.nodeName
				case "TimeStamp"
					tStamp = xmlNode.text
				case "EIndex"
					eIndex = xmlNode.text
                case "Evt"
					eventmsg = xmlNode.text
                case "SensorName"
					sName = xmlNode.text
    		    case "Info"
					info = xmlNode.text
            	case "Value"
					value = xmlNode.text
                case "SensorID"
					sID = xmlNode.text
				case "PointID"
					pID = xmlNode.text
				case "PointIndex"
					pIndex =xmlNode.text
				case "Units"
					units = xmlNode.text
				case "EvtID"
					evtID = xmlNode.text
				case "Alarm"
					alarm = xmlNode.text
			End Select
		Next

		lclStamp = ConvertToLocalTime(tStamp)
		CountIncrVal = 0
		iVal= 0
    next
   xmlString = "<?xml version=""1.0""?>" & vbcrlf
   xmlString = xmlString & "<ErrorList>" & vbcrlf
   xmlString = xmlString & vmcrlf & " <Success>Download Complete</Success>" & vbcrlf
   xmlString = xmlString & "</ErrorList>"

   Response.Write(xmlString)

%>