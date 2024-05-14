<html>
<body>

<%
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

Function GetMonth(utcDateTime)

	GetMonth= DatePart("m", utcDateTime)

End Function

Dim RecTime , CurrentMonth
Dim CurrentIns, IntervalIns

CurrentIns = 0
CurrentMonth = GetMonth(Date)
CurrentYear = DatePart("yyyy",Date)
FirstOfMonth = DateSerial(CurrentYear,CurrentMonth,1)
FirstOfMonthTime = FormatDateTime(FirstOfMonth,VbGeneralDate)
RecMonth = CurrentMonth 


Response.Write("Current Month: " & CurrentMonth & " First of Month Time: " & FirstOfMonthTime)
 
set conn=Server.CreateObject("ADODB.Connection")
'conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "MIUHistory"
set rs = Server.CreateObject("ADODB.recordset")
rs.Open "SELECT MiuName, RecTime,SensorName,IOValue,CountIncr FROM RawHistory where Units='Count' AND CountIncr>0 AND DatePart('m',RecTime)=10 ORDER BY RecTime ASC", conn
%>
<table border="1" width="100%">
<tr>
   <!-- <th>MIUID-MAC</th> -->
   <th>MIUName/Branch</th>
   <th>Time</th>
   <th>SensorName/Location</th>
   <th>All Time Ins</th>
   <th>Ins last 15 minutes</th>
   <th>EM Inside Entrance Total Ins </th>
   <th>EM Outside Entrance Total Ins </th>
   <th>PR Total Ins </th>
  <%do until rs.EOF  %>
  <tr>
  <% 
	RecTime =rs.Fields("RecTime")
	 RecMonth = GetMonth(RecTime)
	 if rs.Fields("MiuName") = "Emmitsburg" AND rs.Fields("SensorName")="Inside Entrance" then
		InsideIntervalIns = rs.Fields("CountIncr")
		InsideCurrentIns = InsideCurrentIns + InsideIntervalIns
	end if
	 if rs.Fields("MiuName") = "Emmitsburg" AND rs.Fields("SensorName")="Outside Entrance" then
		OutsideIntervalIns = rs.Fields("CountIncr")
		OutsideCurrentIns = OutsideCurrentIns + OutsideIntervalIns
	end if
     if InStr(rs.Fields("MiuName"),"Point") AND rs.Fields("SensorName")="1Counter" then
		PRIntervalIns = rs.Fields("CountIncr")
		PRCurrentIns = PRCurrentIns + PRIntervalIns
	end if
  %>
   <% for each x in rs.Fields%>
   <td> <%Response.Write(x.value)%></td> 
  <%next
   rs.MoveNext%>
   
	<td> <%Response.Write(InsideCurrentIns)%> </td> 
	<td> <%Response.Write(OutsideCurrentIns)%> </td> 
    <td> <%Response.Write(PRCurrentIns)%> </td>
</tr>
<%loop
rs.close
conn.close
%>
</table>
</body>
</html>
</body>
</html>