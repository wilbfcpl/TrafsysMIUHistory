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

PRCount = 0
EMOutsideCount = 0
EMInsideCount = 0
RSCount = 0

function GetCounts()

set rs = Server.CreateObject("ADODB.recordset")

	' sql = "SELECT TOP 1 IOValue from RawHistory WHERE MiuName='Point of Rocks' and SensorName='Counter2' and Status = 0 and ucase(Units) like '*COUNT*' and PointIndex=2 and PointType=2 ORDER BY RecTime DESC"
	sql = "SELECT TOP 1 IOValue from RawHistory WHERE MiuName='Point of Rocks' and SensorName='Counter2' and Status = 0 and PointIndex=2 and PointType=2  ORDER BY RecTime DESC"
    rs.open sql, conn 
	RSCount = rs.RecordCount
	if ( IsNumeric (rs.fields("IOValue").value)) then
		PRCount = CLng(rs.fields("IOValue").value)
	end if
	rs.close

	response.write "RSCount: " & RSCount &  " PRCount: " & PRCount & " EMOutsideCount "  & EMOutsideCount & " EMInsideCount " & EMInsideCount & vbcrlf
	'conn.close
end Function 'GetCounts


set conn=Server.CreateObject("ADODB.Connection")
      
strCN = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\\inetpub\\wwwroot\\miu1000\\MIUHistory.mdb;"
conn.open strCN

set objXML = Server.CreateObject("Microsoft.XMLDOM")
objXML.ValidateOnParse = True
objXML.async = False
' Request for real *script launched from MIU-1000* , Canned local file sample.xml for debug
objXML.Load(Request)
'objXML.Load("c:\inetpub\wwwroot\miu1000\sample_single.xml")

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

		lclStamp = ConvertToLocalTime(tStamp)
		CountIncrVal = 0
		iVal= 0

        if IsEmpty(stat) then
            stat = -1
        end if

             if ((ucase(unit) = "COUNT" ) AND (pType=2) AND (pIndex=2) ) then
                  GetCounts()
                  ' IP 10.13.188.251
                  select case ucase(dRecName)
                    case "POINT OF ROCKS"
                         if isNumeric(val) And PRCount>=0 Then
                            'iVal = Cint(val)
                            if iVal > PRCount then
                                CountIncrVal = ( iVal - PRCount  )
                                else
                                 CountIncrVal = 0
                            end if
                           end if
                  end select
           end if
		
           sql = "INSERT INTO RawHistory (MiuID,MiuName,RecTime,Hindex,PointID,SensorID,SensorName,SensorType,PointIndex,PointType,IOValue,iVal,CountIncr,PR,EMInside,EMOutside,Units,Status)"
           sql = sql & " VALUES ('" & dRecID & "','" & dRecName & "','" & lclStamp & "','" & hIndex & "','" & pId & "','" & sID & "','" & sName & "','" & sType & "','"
           sql = sql & pIndex & "','" & pType & "','" & val & "','" & iVal & "','" & CountIncrVal & "','" & PRCount & "','"  & EMInsideCount & "','" & EMOutsideCount & "',' " & unit & "','" & stat & "')"
           conn.Execute(sql & ";")

   Next

   xmlString = "<?xml version=""1.0""?>" & vbcrlf
   xmlString = xmlString & "<ErrorList>" & vbcrlf
   xmlString = xmlString & vmcrlf & " <Success>Download Complete</Success>" & vbcrlf
   xmlString = xmlString & "</ErrorList>"

   Response.Write(xmlString)

    conn.Close
	Set objNode = Nothing
	set objXML = Nothing
	set conn = Nothing
%>
