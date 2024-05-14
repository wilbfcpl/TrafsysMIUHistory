<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<meta http-equiv="X-UA-Compatible" content="ie=edge">
<title>MUI1000</title>
</head>
<body>
<%
	Response.Expires = -1000

	set conn=Server.CreateObject("ADODB.Connection")
	' conn.open = "Driver={Access};Server=(local);Database=MIUHistory;Uid=Trafsys;Pwd=password"
    conn.open ="MIUHistory"
	' Error Handling
	' response.write("<p>")
	' response.write("Open Results")
	' response.write("</p>")
	
	' for each objErr in objConn.Errors
	' response.write("<p>")
	  ' response.write("Description: ")
	  ' response.write(objErr.Description & "<br>")
	  ' response.write("Help context: ")
	  ' response.write(objErr.HelpContext & "<br>")
	  ' response.write("Help file: ")
	  ' response.write(objErr.HelpFile & "<br>")
	  ' response.write("Native error: ")
	  ' response.write(objErr.NativeError & "<br>")
	  ' response.write("Error number: ")
	  ' response.write(objErr.Number & "<br>")
	  ' response.write("Error source: ")
	  ' response.write(objErr.Source & "<br>")
	  ' response.write("SQL state: ")
	  ' response.write(objErr.SQLState & "<br>")
	  ' response.write("</p>")
	' next
	
	set objXML = Server.CreateObject("Microsoft.XMLDOM")
	objXML.ValidateOnParse = True

	' objXML.Load(Request)
	objXML.Load("c:\miu1000\sample.xml")
	If objXML.ParseError.errorCode <> 0 Then
		response.write("<p>")	
		Response.Write("<ErrorList><Error>" & objXML.parseError.reason & "At line: " & objXML.parseError.line & "</Error></ErrorList>")
		response.write("</p>")
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
				case "HIndex"
					hIndex = xmlNode.text
				case "PointID"
					pID = xmlNode.text
				case "SensorID"
					sID = xmlNode.text
				case "SensorName"
					sName = xmlNode.text
				case "SensorType"
					sType = xmlNode.text
				case "PointIndex"
					pIndex =xmlNode.text
				case "PointType"
					pType = xmlNode.text
				case "Value"
					val = xmlNode.text
				case "Units"
					unit = xmlNode.text
				case "Status"
					stat = xmlNode.text
			End Select
		Next
		
		sql = "INSERT INTO RawHistory (MiuID,MiuName,RecTime,Hindex,PointID,SensorID,SensorName,SensorType,PointIndex,PointType,IOValue,Units,Status)"
		sql = sql & " VALUES ('" & dRecID & "','" & dRecName & "','" &tStamp & "','" & hIndex & "','" & pId & "','" & sID & "','" & sName & "','" & sType & "','"
		sql = sql & pIndex & "','" & pType & "','" & val & "','" & unit & "','" & stat & "')"
	    response.write("<p>")	
		Response.Write(sql)
	    response.write("</p>")

		conn.Execute(sql & ";")
	Next
	// xmlString = "<p>" & vbcrlf
	// xmlString = xmlString & "<?xml version=""1.0""?>" & vbcrlf
	xmlString = "<?xml version=""1.0""?>" & vbcrlf
	xmlString = xmlString & "<ErrorList>" & vbcrlf
	xmlString = xmlString & " <Success>Download Complete</Success>" & vbcrlf
	xmlString = xmlString & "</ErrorList>"
	// xmlString = xmlString & "</p>" & vbcrlf

	Response.Write(xmlString)

	conn.Close
	Set objNode = Nothing
	set objXML = Nothing
	set conn = Nothing
%>
</body>
</html>
