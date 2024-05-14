<html>
<body>

<%

dim PRCount
dim EMOutsideCount
dim EMInsideCount 



function GetCounts

PRCount = "TBD"
EMInsideCount = "TBD"
EMOutsideCount = "TBD"


set conn=Server.CreateObject("ADODB.Connection")
'conn.Provider="Microsoft.Jet.OLEDB.4.0"
conn.Open "MIUHistory"
set rs = Server.CreateObject("ADODB.recordset")

sql = "SELECT TOP 1 IOValue from RawHistory WHERE MiuName='Point Of Rocks' AND PointID='0000000064CB2691_2' and SensorName='1Counter' and Units='count' ORDER BY RecTime DESC"
rs.open sql, conn
if ((BOF=False) OR (EOF=False)   ) then
	PRCount = rs.fields("IOValue").value
end if 
rs.close
'Set rs = Nothing

sql = "SELECT TOP 1 IOValue from RawHistory WHERE MiuName='Emmitsburg' AND SensorName='Outside Entrance' and Units='count' ORDER BY RecTime DESC"
rs.open sql,  conn
if ((BOF=False) OR (EOF=False)  ) then
		EMOutsideCount = rs.fields("IOvalue").value
end if
rs.close
'rs = Nothing

sql = "SELECT TOP 1 IOValue from RawHistory WHERE MiuName='Emmitsburg' AND SensorName='Inside Entrance' and Units='count' ORDER BY RecTime DESC"
rs.open sql ,conn
if ( (BOF=False) OR (EOF=False)   )then
	
	EMInsideCount = rs.fields("IOValue").value
	
end if
rs.close 
'rs = Nothing
response.write "PRCount: " + PRCount + " EMOutsideCount " + EMOutsideCount + " EMInsideCount " + EMInsideCount 
conn.close
end Function
GetCounts()
%>
</body>
</html>