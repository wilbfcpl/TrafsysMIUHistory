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

Function GetIPAddress
dim NIC1, Nic, StrIP, CompName

Set NIC1 =     GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

For Each Nic in NIC1

    if Nic.IPEnabled then
        StrIP = Nic.IPAddress(0)
        Set WshNetwork = CreateObject("WScript.Network")
		    if Nic.servicename = "nmvadapter" then
        	'response.write "Nic.name: " & Nic.servicename & " IP Address:  " & StrIP & vbNewLine
			    exit for
		    End if 
    End if
Next
GetIPAddress=StrIP 
end function



set conn=Server.CreateObject("ADODB.Connection")'

strCN = "DRIVER={Microsoft Access Driver (*.mdb)};DBQ=C:\\inetpub\\wwwroot\\miu1000\\MIUHistory.mdb;"

 conn.Open strCN

set rs = Server.CreateObject("ADODB.recordset")
'rs.Open "SELECT TOP 1000 MIUName, RecTime,SensorName,IOValue,CountIncr,PR FROM RawHistory ORDER BY RecTime DESC ", conn
rs.Open "SELECT TOP 1000 RecTime,IOValue,Units, PR FROM RawHistory ORDER BY RecTime DESC ", conn
%>
<table border="1" width="100%">
  <tr><th>Local IP = <a href="http://"><%Response.write GetIPAddress%></a></th></tr>
  <tr><th>Point of Rocks MIU IP Address = <a href="http://10.13.188.251">10.13.188.251</a></th></tr>
</table>
  <table border="1" width="100%">
<tr>
   <!-- <th>MIUID-MAC</th> -->
   <!--th>MIUName/Branch</th>-->
   <th>Time</th>
   <!-- <th>Index</th>
   <th>PointID</th>
   <th>SensorID</th> -->
   <!--th>SensorName/Location</th>-->
   <!-- <th>Sensor Type</th>
   <th>Point Index</th>
   <th>Point Type</th> -->
   <th>IO Value</th>
   <!-- <th>Integer IO</th> -->
   <!--<th>CountIncr</th>-->
   <th>Units</th>
   <th>PR</th>
   
   <!-- <th>Status</th> -->
 </tr>
<%do until rs.EOF%>
  <tr>
  <%for each x in rs.Fields%>
    <td><%Response.Write(x.value)%></td>
  <%next
  rs.MoveNext
%>
  </tr>
<%loop

rs.close
conn.close
%>
</table>

</body>
</html>
