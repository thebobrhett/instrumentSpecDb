<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!--#include file="..\Functions\HitCounter.asp"-->
<html>
<head>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Equipment Inspections</title>
<link rel=STYLESHEET href='http://mogsa4/aksastyle.css' type='text/css'>
<link rel=STYLESHEET href='relief_valve_tds.css' type='text/css'>
<style type="text/css">
caption { font-weight:bold; font-size:14; color:darkblue}
th {background-color:powderblue; font-weight:bold; font-size:12; color:black}
td { text-align:left; }
a  { text-decoration:underline; }
.center {text-align:center; }
</style> 
</head>
<body>

<div id='title'>
Relief Valve - Technical Data Sheet
</div>

<div id='logo'>
<img src='http://mogsa4/images/akgrouplogosmallest.gif' />
</div>

<p align='center'>
<%
'*************
'Bob Rhett - Tuesday, September 29, 2009
'  Created
'*************

'on error resume next

dim objEQdb
dim objEQrs
dim formatted_ts
dim strSQL
dim HitCounts

formatted_ts = now()
formatted_ts = cstr(year(formatted_ts)) & "-" & cstr(month(formatted_ts)) & "-" & cstr(day(formatted_ts)) & " " & cstr(hour(formatted_ts)) & ":" & cstr(minute(formatted_ts)) & ":" & cstr(second(formatted_ts))

'Set/get hit counts.
HitCounts = HitCounter("threadline_home")

set objEQdb = CreateObject("adodb.connection")
objEQdb.open = "driver={MySQL ODBC 3.51 Driver};option=16387;server=richmond.aksa.local;user=rootb;password=spandex;DATABASE=equipment;"
set objEQrs = CreateObject("adodb.recordset")

strSQL = "select * from relief_valve"
objEQrs.open strSQL, objEQdb
response.write "<div id='header-left'>"
response.write "<table>"
response.write "<tr><td>MMIS No: </td><td>" & objEQrs("mmis_no") & "</td></tr>"
response.write "<tr><td>Assembly: </td><td>" & objEQrs("assembly") & "</td></tr>"
response.write "<tr><td>Equipment Name: </td><td>" & objEQrs("eq_name") & "</td></tr>"
response.write "<tr><td>Equipment No: </td><td>" & objEQrs("eq_no") & "</td></tr>"
response.write "<tr><td>Asset No: </td><td>" & objEQrs("asset_no") & "</td></tr>"
response.write "</table>"
response.write "</div>"

response.write "<div id='header-right'>"
response.write "<table>"
response.write "<tr><td>Plant Location: </td><td>" & objEQrs("plant_location") & "</td></tr>"
response.write "<tr><td>Zone: </td><td>" & objEQrs("zone") & "</td></tr>"
response.write "<tr><td>Area: </td><td>" & objEQrs("area") & "</td></tr>"
response.write "<tr><td>Cost Center: </td><td>" & objEQrs("cc") & "</td></tr>"
response.write "<tr><td>Division: </td><td>" & objEQrs("division") & "</td></tr>"
response.write "</table>"
response.write "</div>"

objEQrs.close

set objEQrs = nothing
objEQdb.close
set objEQdb = nothing
%>
</p>
</body>
</html>
