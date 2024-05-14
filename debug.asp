<%
	Response.Expires = -1000
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

Function GetIPAddress
dim NIC1, Nic, StrIP, CompName

Set NIC1 =     GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

For Each Nic in NIC1

    if Nic.IPEnabled then
        StrIP = Nic.IPAddress(0)

        Set WshNetwork = CreateObject("WScript.Network")
		if Nic.servicename = "nmvadapter" then
        	response.write "Nic.name: " & Nic.servicename & " IP Address:  " & StrIP & vbNewLine
			exit for
		End if 
    End if
Next
end function


PRCount = 0
RSCount = 0

function GetCounts



' set conn=Server.CreateObject("ADODB.Connection")
' 'conn.Provider="Microsoft.Jet.OLEDB.4.0"
' conn.Open "MIUHistory"
set rs = Server.CreateObject("ADODB.recordset")

sql = "SELECT TOP 1 IOValue from RawHistory WHERE MiuName='Point of Rocks' and SensorName='1Counter' and Status = 0 and Units='count' and PointIndex=2 and PointType=2 ORDER BY RecTime DESC"
rs.open sql, conn 
RSCount = rs.RecordCount
if ( IsNumeric (rs.fields("IOValue").value)) then
	PRCount = Cint(rs.fields("IOValue").value)
end if
rs.close
Set rs = Nothing

response.write "RSCount: " & RSCount &  " PRCount: " & PRCount
'conn.close
end Function 'GetCounts

' Main routine part of the script'

GetIPAddress

set conn=Server.CreateObject("ADODB.Connection")
      '   conn.Open "Provider=Microsoft.Jet.OLEDB.4.0;data source=c:/inetpub/wwwroot/miu1000/MIUHistory.mdb;userID=DefaultAppPool;password=;"
	 'conn.open = "Driver={Access};Server=(local);Database=MIUHistory;Uid=DefaultAppPool"
    	 ' conn.open "MIUHistory","IIS AppPool\DefaultAppPool",""
	'  conn.open "MIUHistory","IUSR",""	
	'conn.open ".\MIUHistory.mdb"

	strCN = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\\inetpub\\wwwroot\\miu1000\\MIUHistory.mdb;"
    conn.open strCN

	set objXML = Server.CreateObject("Microsoft.XMLDOM")
	objXML.ValidateOnParse = True

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
		
		if (status= 0) and (unit = "count" ) AND (pType=2) and (pIndex=2) then
			
		 GetCounts()
		
			select case dRecName
			case "Point of Rocks"
				'if pID="0000000064CB2691" and isNumeric(val) And PRCount>=0 Then
				if isNumeric(val) And PRCount>=0 Then
				iVal = Cint(val)	
				if iVal > PRCount then
					CountIncrVal = ( iVal - PRCount  )
					else
					CountIncrVal = 0
				 end if 
				end if
			
			
			end select
		end if 'if (unit = "count" ) AND (pType=2) and (pIndex=2)then
		
		sql = "INSERT INTO RawHistory (MiuID,MiuName,RecTime,Hindex,PointID,SensorID,SensorName,SensorType,PointIndex,PointType,IOValue,iVal,CountIncr,PR,Units,Status)"
		sql = sql & " VALUES ('" & dRecID & "','" & dRecName & "','" & lclStamp & "','" & hIndex & "','" & pId & "','" & sID & "','" & sName & "','" & sType & "','"
		sql = sql & pIndex & "','" & pType & "','" & val & "','" & iVal & "','" & CountIncrVal & "','" & PRCount & "','"  & unit & "','" & stat & "')"
	    ' response.write("<p>")	
            ' Response.Write(sql)
	    ' response.write("</p>")

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
