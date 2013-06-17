<html>
<head>
<title>Test Ascii</title>
<!--#include file="EquipmentFunctions.asp"-->
</head>
<body>
<%
Dim sqlString
Dim cn
Dim rs
Dim temp
Dim count

'Define the ado connection and recordset objects.
set cn = CreateObject("adodb.connection")
cn.Open = DBString
set rs = CreateObject("adodb.recordset")

sqlString = "SELECT instr_name FROM instruments WHERE instr_id=171615"
Set rs = cn.Execute(sqlString)
temp = rs(0)
For count = 1 To Len(temp)
	Response.Write count & " = " & Asc(Mid(temp,count,1)) & "<br />"
Next

Response.Write "<br />"

sqlString = "SELECT instr_name FROM instruments WHERE instr_id=36224"
Set rs = cn.Execute(sqlString)
temp = rs(0)
For count = 1 To Len(temp)
	Response.Write count & " = " & Asc(Mid(temp,count,1)) & "<br />"
Next

rs.Close
Set rs = Nothing
cn.Close
Set cn = Nothing
%>
</body>
</html>
