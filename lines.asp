<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="javascript">
function doDelete(id) {
	if (confirm("Are you sure you want to delete record number "+id+"?")) {
		document.form1.action="adminaction.asp?action=delete&RECORD="+id;
		document.form1.submit()
	}
}
function openhelp() {
 window.open("Instrument Spec Database Administrators Guide.doc","userguide");
}
function reloadPage() {
 document.form1.action="lines.asp";
 document.form1.submit()
}
function findRecord(id) {
 var newid = id
 document.form1.action="lines.asp?lastnum="+newid;
 document.form1.submit()
}
</script>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Administer Lines</title>
<link rel=STYLESHEET href='equipmentstyle.css' type='text/css'>
<style>
  div {font-family:verdana;}
  input {font-family:verdana;
		font-size:8pt;
		background-color:#DBF5F5}
  select {font-family:verdana;
		font-size:8pt;
		background-color:#DBF5F5}
  textarea {font-family:verdana;
		font-size:8pt;
		background-color:#DBF5F5}
<!--
  a     { text-decoration:underline; }
-->
</style> 
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="EquipmentFunctions.asp"-->
</head>

<%
'*************
' Revision History
' 
' Keith Brooks - Tuesday, April 27, 2010
'   Creation.
'*************

'on error resume next

dim cn
dim rs
Dim rs2
dim recordid
Dim firstPass
Dim firstNum
Dim lastNum
Dim maxNum
Dim minNum
Dim currentuser
Dim access
Dim recid
Dim sqlString
Dim limitnum
Dim sortkey
Dim sortdir

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("equipment", "lines", currentuser)
If access <> "none" Then

	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")
	set rs2 = CreateObject("adodb.recordset")

	If session("err") <> "" And session("err") <> "NONE" Then
	  Response.Write "<body onload='document.form1." & session("err") & ".focus();'>"
	ElseIf session("focus") <> "" And session("focus") <> "NONE" Then
	  Response.Write "<body onload='document.form1." & session("focus") & ".focus();'>"
	  session("focus") = "NONE"
	Else
	  Response.Write "<body>"
	End If

	If request("record_id") <> "" Then
	  If IsNumeric(request("record_id")) Then
	    recordid = request("record_id")
	  Else
	    recordid = 0
	  End If
	Else
	  recordid = 0
	End If
	If request("sort") <> "" Then
		sortkey = request("sort")
	Else
		sortkey = "line_id"
	End If
	If request("direction") <> "" Then
		sortdir = request("direction")
	Else
		sortdir = "ASC"
	End If
	If Request("limit") <> "" Then
		limitnum = Request("limit")
	Else
		limitnum = "20"
	End If

	response.write "<table ID='headertable' width='100%'>"
	response.write "<tr>"
	response.write "<td ID='headertd' style='width:30%;text-align:left;vertical-align:top'><a href='adminmenu.asp' title='Open the administration main menu'>Menu</a></td>"
	response.write "<td ID='headertd' style='width:40%;text-align:center;vertical-align:center'><h1/>Edit Lines</td>"
	response.write "<td ID='headertd' style='width:30%;text-align:right;vertical-align:top'><a href='#' onClick='openhelp();return false;' title='Open the Admin Guide'>Help</a></td>"
	response.write "</tr>"

	response.Write "<form action='adminaction.asp' id='form1' method='post' name='form1'>"

	response.write "<tr>"
	response.write "<td id='headertd' style='text-align:left'>Find Line:"
	Response.Write "<select name='find' id='find' tabindex='1' onchange='findRecord(this.value);'>"
	sqlString = "SELECT line_id,line_num FROM process_lines ORDER BY line_num"
	Set rs = cn.Execute(sqlString)
	If Not rs.BOF Then
		rs.MoveFirst
		Response.Write "<option value=''> "
		Do While Not rs.EOF
			If IsNumeric(Request("find")) Then
				If CLng(Request("find")) = rs("line_id") Then
					Response.Write "<option value='" & rs("line_id") & "' selected>" & rs("line_num")
				Else
					Response.Write "<option value='" & rs("line_id") & "'>" & rs("line_num")
				End If
			Else
				Response.Write "<option value='" & rs("line_id") & "'>" & rs("line_num")
			End If
			rs.MoveNext
		Loop
	End If
	rs.Close
	If access = "write" Or access = "delete" Then
		response.write "<td ID='headertd' style='text-align:center'><a href='lines.asp?record_id=-1&sort=" & sortkey & "&direction=" & sortdir & "&limit=" & limitnum & "' tabindex='2' title='Add a new line record'>Add new record</a></td>"
	Else
		Response.Write "<td id='headertd'>&nbsp;</td>"
	End If
	response.write "<td id='headertd'>Records to display:"
	Response.Write "<select name='limit' id='limit' tabindex='3' onchange='reloadPage();'>"
	If limitnum = "All" Then
		Response.Write "<option value='All' selected>All"
	Else
		Response.Write "<option value='All'>All"
	End If
	If limitnum = "20" Then
		Response.Write "<option value='20' selected>20"
	Else
		Response.Write "<option value='20'>20"
	End If
	If limitnum = "100" Then
		Response.Write "<option value='100' selected>100"
	Else
		Response.Write "<option value='100'>100"
	End If
	If limitnum = "1000" Then
		Response.Write "<option value='1000' selected>1000"
	Else
		Response.Write "<option value='1000'>1000"
	End If
	Response.Write "</select></td>"
	response.write "</tr>"
	response.write "</table>"

	Response.Write "<br />"

	response.write "<input type='hidden' name='RECORD' value='" & recordid & "'>"
	response.write "<input type='hidden' name='SORT' value='" & sortkey & "'>"
	response.write "<input type='hidden' name='DIRECTION' value='" & sortdir & "'>"

'	If IsNumeric(limitNum) Then
		'Get the max record number.
'		sqlString = "SELECT MAX(line_id) FROM process_lines"
'		Set rs = cn.Execute(sqlString)
'		If Not rs.BOF Then
'			rs.MoveFirst
'			maxNum = rs(0)
'		Else
'			maxNum = ""
'		End If
'		rs.Close

		'Get the min record number.
'		sqlString = "SELECT MIN(line_id) FROM process_lines"
'		Set rs = cn.Execute(sqlString)
'		If Not rs.BOF Then
'			rs.MoveFirst
'			minNum = rs(0)
'		Else
'			minNum = ""
'		End If
'		rs.Close

		'If the user has clicked the "Prev" link, find the record number to start with.
'		If Request("firstnum") <> "" Then
'			sqlString = "SELECT line_id FROM process_lines WHERE line_id > " & Request("firstnum") & " ORDER BY line_id ASC LIMIT " & limitNum
'			Set rs = cn.Execute(sqlString)
'			If Not rs.BOF Then
'				rs.MoveFirst
'				Do While Not rs.EOF
'					firstNum = rs(0)
'					rs.MoveNext
'				Loop
'			Else
'				firstNum = Request("firstnum")
'			End If
'			rs.Close
'		End If
		
'	End If
	
	'Read the data.
	sqlString = "SELECT process_lines.*,plant_area_name AS plant_area," & _
			"line_type_name AS line_type,pipe_mat_name AS pipe_mat," & _
			"dwg_name AS dwg,pipe_class_name AS pipe_class " & _
			"FROM ((((process_lines LEFT JOIN plant_areas " & _
			"ON process_lines.plant_area_id=plant_areas.plant_area_id) " & _
			"LEFT JOIN line_types ON process_lines.line_type_id=line_types.line_type_id) " & _
			"LEFT JOIN pipe_materials ON process_lines.pipe_orif_mat_id=pipe_materials.pipe_mat_id) " & _
			"LEFT JOIN drawings ON process_lines.dwg_id=drawings.dwg_id) " & _
			"LEFT JOIN pipe_classes ON process_lines.pipe_class_id=pipe_classes.pipe_class_id "
	If IsNumeric(limitNum) Then
		If request("lastnum") <> "" Then
			sqlString = sqlString & "WHERE line_id = " & Request("lastnum") & _
					" ORDER BY " & sortkey & " " & sortdir & " LIMIT " & limitNum
'		ElseIf Request("firstnum") <> "" Then
''			sqlString = sqlString & "WHERE line_id < " & CStr(CLng(request("firstnum")) + CInt(limitNum) + 1) & _
'			sqlString = sqlString & "WHERE line_id <= " & firstNum & _
'					" ORDER BY " & sortkey & " " & sortdir & " LIMIT " & limitNum
		Else
			sqlString = sqlString & "ORDER BY " & sortkey & " " & sortdir & " LIMIT " & limitNum
		End If
	Else
		sqlString = sqlString & "ORDER BY ORDER BY " & sortkey & " " & sortdir
	End If
	set rs = cn.Execute(sqlString)

	If Not rs.BOF Then
		rs.MoveFirst
	End If
	  
	'If recordid<0, the user has selected "Add new record" so insert a blank data entry line
	'at the top of the form.
	If access = "write" Or access = "delete" Then
		If recordid < 0 Then

			'Draw the first header row.
			response.Write "<table width='100%'>"
			response.Write "<tr>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_id&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Line ID&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_id&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=plant_area&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Plant Area&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=plant_area&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_num&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Line Number&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_num&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_size&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Line Size&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_size&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=wall_thick&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Wall<br />Thickness&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=wall_thick&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=rtg&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Rating&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=rtg&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_type&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Line Type&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_type&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_sched&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Line<br />Schedule&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_sched&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_internal_dia&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Internal<br />Diameter&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_internal_dia&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "</tr>"

			Response.Write "<tr>"
			response.write "<td id='mediumtd'>&nbsp;</td>"

			'Dropdown for plant area.
			If session("err") = "plant_area_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='plant_area_id' tabindex='4'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='plant_area_id' tabindex='4'>"
			End If
			sqlString = "SELECT plant_area_id,plant_area_name FROM plant_areas ORDER BY plant_area_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("plant_area_id") <> "" Then
						If CLng(Request("plant_area_id")) = rs2("plant_area_id") Then
							response.write "<option value='" & rs2("plant_area_id") & "' selected>" & rs2("plant_area_name")
						Else
							response.write "<option value='" & rs2("plant_area_id") & "'>" & rs2("plant_area_name")
						End If
					Else
						If Session("plant_area_id") <> "" Then
							If CLng(session("plant_area_id")) = rs2("plant_area_id") Then
								response.write "<option value='" & rs2("plant_area_id") & "' selected>" & rs2("plant_area_name")
							Else
								response.write "<option value='" & rs2("plant_area_id") & "'>" & rs2("plant_area_name")
							End If
						Else
							response.write "<option value='" & rs2("plant_area_id") & "'>" & rs2("plant_area_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			If session("err") = "line_num" Then
				If session("line_num") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_num' size='20' tabindex='5' value='" & session("line_num") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_num' size='20' tabindex='5' value=''></td>"
				End If
			Else
				If session("line_num") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='line_num' size='20' tabindex='5' value='" & session("line_num") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='line_num' size='20' tabindex='5' value=''></td>"
				End If
			End If

			If session("err") = "line_size" Then
				If session("line_size") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_size' size='5' tabindex='6' value='" & session("line_size") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_size' size='5' tabindex='6' value=''></td>"
				End If
			Else
				If session("line_size") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='line_size' size='5' tabindex='6' value='" & session("line_size") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='line_size' size='5' tabindex='6' value=''></td>"
				End If
			End If

			If session("err") = "wall_thick" Then
				If session("wall_thick") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='wall_thick' size='5' tabindex='7' value='" & session("wall_thick") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='wall_thick' size='5' tabindex='7' value=''></td>"
				End If
			Else
				If session("wall_thick") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='wall_thick' size='5' tabindex='7' value='" & session("wall_thick") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='wall_thick' size='5' tabindex='7' value=''></td>"
				End If
			End If

			If session("err") = "rtg" Then
				If session("rtg") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='rtg' size='5' tabindex='8' value='" & session("rtg") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='rtg' size='5' tabindex='8' value=''></td>"
				End If
			Else
				If session("rtg") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='rtg' size='5' tabindex='8' value='" & session("rtg") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='rtg' size='5' tabindex='8' value=''></td>"
				End If
			End If

			'Dropdown for line type.
			If session("err") = "line_type_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='line_type_id' tabindex='9'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='line_type_id' tabindex='9'>"
			End If
			sqlString = "SELECT line_type_id,line_type_name FROM line_types ORDER BY line_type_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("line_type_id") <> "" Then
						If CLng(Request("line_type_id")) = rs2("line_type_id") Then
							response.write "<option value='" & rs2("line_type_id") & "' selected>" & rs2("line_type_name")
						Else
							response.write "<option value='" & rs2("line_type_id") & "'>" & rs2("line_type_name")
						End If
					Else
						If Session("line_type_id") <> "" Then
							If CLng(session("line_type_id")) = rs2("line_type_id") Then
								response.write "<option value='" & rs2("line_type_id") & "' selected>" & rs2("line_type_name")
							Else
								response.write "<option value='" & rs2("line_type_id") & "'>" & rs2("line_type_name")
							End If
						Else
							response.write "<option value='" & rs2("line_type_id") & "'>" & rs2("line_type_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			If session("err") = "line_sched" Then
				If session("line_sched") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_sched' size='5' tabindex='10' value='" & session("line_sched") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_sched' size='5' tabindex='10' value=''></td>"
				End If
			Else
				If session("line_sched") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='line_sched' size='5' tabindex='10' value='" & session("line_sched") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='line_sched' size='5' tabindex='10' value=''></td>"
				End If
			End If

			If session("err") = "line_internal_dia" Then
				If session("line_internal_dia") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_internal_dia' size='5' tabindex='11' value='" & session("line_internal_dia") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_internal_dia' size='5' tabindex='11' value=''></td>"
				End If
			Else
				If session("line_internal_dia") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='line_internal_dia' size='5' tabindex='11' value='" & session("line_internal_dia") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='line_internal_dia' size='5' tabindex='11' value=''></td>"
				End If
			End If
			Response.Write "</tr>"
			Response.Write "</table>"

			'Draw the second header row.
			response.Write "<table width='100%'>"
			response.Write "<tr>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_std_id&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Pipe<br />Std ID&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_std_id&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=ansi_din&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;ANSI<br />DIN&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=ansi_din&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_size&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Pipe Size&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_size&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_mat&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Pipe Material&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_mat&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=stream_num&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Stream&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=stream_num&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=dwg&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Drawing&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=dwg&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_class&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Pipe Class&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_class&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			If access = "write" Or access = "delete" Then
 			  response.Write "<th id='mediumth'>&nbsp;</th>"
 			End If
			If access = "delete" Or recordid < 0 Then
 			  response.Write "<th id='mediumth'>&nbsp;</th>"
 			End If
			response.Write "</tr>"

			If session("err") = "pipe_std_id" Then
				If session("pipe_std_id") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='pipe_std_id' size='5' tabindex='12' value='" & session("pipe_std_id") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='pipe_std_id' size='5' tabindex='12' value=''></td>"
				End If
			Else
				If session("pipe_std_id") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='pipe_std_id' size='5' tabindex='12' value='" & session("pipe_std_id") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='pipe_std_id' size='5' tabindex='12' value=''></td>"
				End If
			End If

			If session("err") = "ansi_din" Then
				If session("ansi_din") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='ansi_din' size='5' tabindex='13' value='" & session("ansi_din") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='ansi_din' size='5' tabindex='13' value=''></td>"
				End If
			Else
				If session("ansi_din") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='ansi_din' size='5' tabindex='13' value='" & session("ansi_din") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='ansi_din' size='5' tabindex='13' value=''></td>"
				End If
			End If

			If session("err") = "pipe_size" Then
				If session("pipe_size") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='pipe_size' size='5' tabindex='14' value='" & session("pipe_size") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='pipe_size' size='5' tabindex='14' value=''></td>"
				End If
			Else
				If session("pipe_size") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='pipe_size' size='5' tabindex='14' value='" & session("pipe_size") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='pipe_size' size='5' tabindex='14' value=''></td>"
				End If
			End If

			'Dropdown for pipe material.
			If session("err") = "pipe_orif_mat_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='pipe_orif_mat_id' tabindex='15'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='pipe_orif_mat_id' tabindex='15'>"
			End If
			sqlString = "SELECT pipe_mat_id,pipe_mat_name FROM pipe_materials ORDER BY pipe_mat_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("pipe_orif_mat_id") <> "" Then
						If CLng(Request("pipe_orif_mat_id")) = rs2("pipe_mat_id") Then
							response.write "<option value='" & rs2("pipe_mat_id") & "' selected>" & rs2("pipe_mat_name")
						Else
							response.write "<option value='" & rs2("pipe_mat_id") & "'>" & rs2("pipe_mat_name")
						End If
					Else
						If Session("pipe_orif_mat_id") <> "" Then
							If CLng(session("pipe_orif_mat_id")) = rs2("pipe_mat_id") Then
								response.write "<option value='" & rs2("pipe_mat_id") & "' selected>" & rs2("pipe_mat_name")
							Else
								response.write "<option value='" & rs2("pipe_mat_id") & "'>" & rs2("pipe_mat_name")
							End If
						Else
							response.write "<option value='" & rs2("pipe_mat_id") & "'>" & rs2("pipe_mat_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			If session("err") = "stream_num" Then
				If session("stream_num") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='stream_num' size='20' tabindex='16' value='" & session("stream_num") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='stream_num' size='20' tabindex='16' value=''></td>"
				End If
			Else
				If session("stream_num") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='stream_num' size='20' tabindex='16' value='" & session("stream_num") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='stream_num' size='20' tabindex='16' value=''></td>"
				End If
			End If

			'Dropdown for drawings.
			If session("err") = "dwg_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='dwg_id' tabindex='17'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='dwg_id' tabindex='17'>"
			End If
			sqlString = "SELECT dwg_id,dwg_name FROM drawings WHERE dwg_name is not null and dwg_type_id=1 ORDER BY dwg_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("dwg_id") <> "" Then
						If CLng(Request("dwg_id")) = rs2("dwg_id") Then
							response.write "<option value='" & rs2("dwg_id") & "' selected>" & rs2("dwg_name")
						Else
							response.write "<option value='" & rs2("dwg_id") & "'>" & rs2("dwg_name")
						End If
					Else
						If Session("dwg_id") <> "" Then
							If CLng(session("dwg_id")) = rs2("dwg_id") Then
								response.write "<option value='" & rs2("dwg_id") & "' selected>" & rs2("dwg_name")
							Else
								response.write "<option value='" & rs2("dwg_id") & "'>" & rs2("dwg_name")
							End If
						Else
							response.write "<option value='" & rs2("dwg_id") & "'>" & rs2("dwg_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			'Dropdown for pipe_class.
			If session("err") = "pipe_class_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='pipe_class_id' tabindex='18'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='pipe_class_id' tabindex='18'>"
			End If
			sqlString = "SELECT pipe_class_id,pipe_class_name FROM pipe_classes ORDER BY pipe_class_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("pipe_class_id") <> "" Then
						If CLng(Request("pipe_class_id")) = rs2("pipe_class_id") Then
							response.write "<option value='" & rs2("pipe_class_id") & "' selected>" & rs2("pipe_class_name")
						Else
							response.write "<option value='" & rs2("pipe_class_id") & "'>" & rs2("pipe_class_name")
						End If
					Else
						If Session("pipe_class_id") <> "" Then
							If CLng(session("pipe_class_id")) = rs2("pipe_class_id") Then
								response.write "<option value='" & rs2("pipe_class_id") & "' selected>" & rs2("pipe_class_name")
							Else
								response.write "<option value='" & rs2("pipe_class_id") & "'>" & rs2("pipe_class_name")
							End If
						Else
							response.write "<option value='" & rs2("pipe_class_id") & "'>" & rs2("pipe_class_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' tabindex='19' title='Apply changes to this record'></td>"

			Response.Write "<td id='mediumtd' style='text-align:center'><a href='lines.asp?limit=" & limitnum & "' tabindex='20' title='Cancel changes to this record'>Cancel</a></td>"

			Response.Write "</tr>"
			Response.Write "</table>"
			
			Response.Write "<table style='width:100%'>"
			Response.Write "<tr><td style='background-color:blue;font-size:1px;padding-top:0px;padding-bottom:0px;width:100%'>&nbsp;</td></tr>"
			Response.Write "</table>"
		End If
	End If

	firstPass = True
	Do While Not rs.EOF

		'Store the first and last record numbers.
		If firstPass Then
			firstNum = rs("line_id")
		End If
		firstPass = False
		lastNum = rs("line_id")

		'Draw the first header row.
			response.Write "<table width='100%'>"
			response.Write "<tr>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_id&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Line ID&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_id&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=plant_area&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Plant Area&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=plant_area&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_num&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Line Number&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_num&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_size&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Line Size&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_size&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=wall_thick&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Wall<br />Thickness&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=wall_thick&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=rtg&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Rating&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=rtg&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_type&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Line Type&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_type&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_sched&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Line<br />Schedule&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_sched&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_internal_dia&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Internal<br />Diameter&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=line_internal_dia&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "</tr>"

		Response.Write "<tr>"
		If CLng(rs("line_id")) = CLng(recordid) Then
			'Draw the first data entry line
			response.write "<td id='mediumtd'>" & rs("line_id") & "</td>"

			'Dropdown for plant area.
			If session("err") = "plant_area_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='plant_area_id' tabindex='4'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='plant_area_id' tabindex='4'>"
			End If
			sqlString = "SELECT plant_area_id,plant_area_name FROM plant_areas ORDER BY plant_area_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("plant_area_id") <> "" Then
						If CLng(Request("plant_area_id")) = rs2("plant_area_id") Then
							response.write "<option value='" & rs2("plant_area_id") & "' selected>" & rs2("plant_area_name")
						Else
							response.write "<option value='" & rs2("plant_area_id") & "'>" & rs2("plant_area_name")
						End If
					Else
						If rs("plant_area") = rs2("plant_area_name") Then
							response.write "<option value='" & rs2("plant_area_id") & "' selected>" & rs2("plant_area_name")
						Else
							response.write "<option value='" & rs2("plant_area_id") & "'>" & rs2("plant_area_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			If session("err") = "line_num" Then
				If session("line_num") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_num' size='20' tabindex='5' value='" & session("line_num") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_num' size='20' tabindex='5' value='" & rs("line_num") & "'></td>"
				End If
			Else
				If session("line_num") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='line_num' size='20' tabindex='5' value='" & session("line_num") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='line_num' size='20' tabindex='5' value='" & rs("line_num") & "'></td>"
				End If
			End If

			If session("err") = "line_size" Then
				If session("line_size") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_size' size='5' tabindex='6' value='" & session("line_size") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_size' size='5' tabindex='6' value='" & rs("line_size") & "'></td>"
				End If
			Else
				If session("line_size") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='line_size' size='5' tabindex='6' value='" & session("line_size") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='line_size' size='5' tabindex='6' value='" & rs("line_size") & "'></td>"
				End If
			End If

			If session("err") = "wall_thick" Then
				If session("wall_thick") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='wall_thick' size='5' tabindex='7' value='" & session("wall_thick") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='wall_thick' size='5' tabindex='7' value='" & rs("wall_thick") & "'></td>"
				End If
			Else
				If session("wall_thick") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='wall_thick' size='5' tabindex='7' value='" & session("wall_thick") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='wall_thick' size='5' tabindex='7' value='" & rs("wall_thick") & "'></td>"
				End If
			End If

			If session("err") = "rtg" Then
				If session("rtg") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='rtg' size='5' tabindex='8' value='" & session("rtg") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='rtg' size='5' tabindex='8' value='" & rs("rtg") & "'></td>"
				End If
			Else
				If session("rtg") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='rtg' size='5' tabindex='8' value='" & session("rtg") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='rtg' size='5' tabindex='8' value='" & rs("rtg") & "'></td>"
				End If
			End If

			'Dropdown for line type.
			If session("err") = "line_type_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='line_type_id' tabindex='9'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='line_type_id' tabindex='9'>"
			End If
			sqlString = "SELECT line_type_id,line_type_name FROM line_types ORDER BY line_type_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("line_type_id") <> "" Then
						If CLng(Request("line_type_id")) = rs2("line_type_id") Then
							response.write "<option value='" & rs2("line_type_id") & "' selected>" & rs2("line_type_name")
						Else
							response.write "<option value='" & rs2("line_type_id") & "'>" & rs2("line_type_name")
						End If
					Else
						If rs("line_type") = rs2("line_type_name") Then
							response.write "<option value='" & rs2("line_type_id") & "' selected>" & rs2("line_type_name")
						Else
							response.write "<option value='" & rs2("line_type_id") & "'>" & rs2("line_type_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			If session("err") = "line_sched" Then
				If session("line_sched") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_sched' size='5' tabindex='10' value='" & session("line_sched") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_sched' size='5' tabindex='10' value='" & rs("line_sched") & "'></td>"
				End If
			Else
				If session("line_sched") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='line_sched' size='5' tabindex='10' value='" & session("line_sched") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='line_sched' size='5' tabindex='10' value='" & rs("line_sched") & "'></td>"
				End If
			End If

			If session("err") = "line_internal_dia" Then
				If session("line_internal_dia") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_internal_dia' size='5' tabindex='11' value='" & session("line_internal_dia") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='line_internal_dia' size='5' tabindex='11' value='" & rs("line_internal_dia") & "'></td>"
				End If
			Else
				If session("line_internal_dia") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='line_internal_dia' size='5' tabindex='11' value='" & session("line_internal_dia") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='line_internal_dia' size='5' tabindex='11' value='" & rs("line_internal_dia") & "'></td>"
				End If
			End If

		Else
			'Draw the first history row.
			response.write "<td id='mediumtd'>" & rs("line_id") & "</td>"
			If rs("plant_area") <> "" Then
				response.write "<td id='mediumtd'>" & rs("plant_area") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("line_num")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("line_num") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("line_size")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("line_size") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("wall_thick")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("wall_thick") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("rtg")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("rtg") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("line_type") <> "" Then
				response.write "<td id='mediumtd'>" & rs("line_type") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("line_sched")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("line_sched") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("line_internal_dia")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("line_internal_dia") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
		End If
		response.write "</tr>"
		Response.Write "</table>"

		'Draw the second header row.
			response.Write "<table width='100%'>"
			response.Write "<tr>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_std_id&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Pipe<br />Std ID&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_std_id&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=ansi_din&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;ANSI<br />DIN&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=ansi_din&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_size&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Pipe Size&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_size&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_mat&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Pipe Material&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_mat&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=stream_num&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Stream&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=stream_num&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=dwg&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Drawing&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=dwg&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_class&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Pipe Class&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='lines.asp?sort=pipe_class&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		If access = "write" Or access = "delete" Then
 		  response.Write "<th id='mediumth'>&nbsp;</th>"
 		End If
		If access = "delete" Or recordid < 0 Then
 		  response.Write "<th id='mediumth'>&nbsp;</th>"
 		End If
		response.Write "</tr>"

		Response.Write "<tr>"
		If CLng(rs("line_id")) = CLng(recordid) Then
			'Draw the second data entry line
			If session("err") = "pipe_std_id" Then
				If session("pipe_std_id") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='pipe_std_id' size='5' tabindex='12' value='" & session("pipe_std_id") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='pipe_std_id' size='5' tabindex='12' value='" & rs("pipe_std_id") & "'></td>"
				End If
			Else
				If session("pipe_std_id") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='pipe_std_id' size='5' tabindex='12' value='" & session("pipe_std_id") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='pipe_std_id' size='5' tabindex='12' value='" & rs("pipe_std_id") & "'></td>"
				End If
			End If

			If session("err") = "ansi_din" Then
				If session("ansi_din") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='ansi_din' size='5' tabindex='13' value='" & session("ansi_din") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='ansi_din' size='5' tabindex='13' value='" & rs("ansi_din") & "'></td>"
				End If
			Else
				If session("ansi_din") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='ansi_din' size='5' tabindex='13' value='" & session("ansi_din") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='ansi_din' size='5' tabindex='13' value='" & rs("ansi_din") & "'></td>"
				End If
			End If

			If session("err") = "pipe_size" Then
				If session("pipe_size") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='pipe_size' size='5' tabindex='14' value='" & session("pipe_size") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='pipe_size' size='5' tabindex='14' value='" & rs("pipe_size") & "'></td>"
				End If
			Else
				If session("pipe_size") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='pipe_size' size='5' tabindex='14' value='" & session("pipe_size") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='pipe_size' size='5' tabindex='14' value='" & rs("pipe_size") & "'></td>"
				End If
			End If

			'Dropdown for pipe material.
			If session("err") = "pipe_orif_mat_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='pipe_orif_mat_id' tabindex='15'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='pipe_orif_mat_id' tabindex='15'>"
			End If
			sqlString = "SELECT pipe_mat_id,pipe_mat_name FROM pipe_materials ORDER BY pipe_mat_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("pipe_orif_mat_id") <> "" Then
						If CLng(Request("pipe_orif_mat_id")) = rs2("pipe_mat_id") Then
							response.write "<option value='" & rs2("pipe_mat_id") & "' selected>" & rs2("pipe_mat_name")
						Else
							response.write "<option value='" & rs2("pipe_mat_id") & "'>" & rs2("pipe_mat_name")
						End If
					Else
						If rs("pipe_mat") = rs2("pipe_mat_name") Then
							response.write "<option value='" & rs2("pipe_mat_id") & "' selected>" & rs2("pipe_mat_name")
						Else
							response.write "<option value='" & rs2("pipe_mat_id") & "'>" & rs2("pipe_mat_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			If session("err") = "stream_num" Then
				If session("stream_num") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='stream_num' size='20' tabindex='16' value='" & session("stream_num") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='stream_num' size='20' tabindex='16' value='" & rs("stream_num") & "'></td>"
				End If
			Else
				If session("stream_num") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='stream_num' size='20' tabindex='16' value='" & session("stream_num") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='stream_num' size='20' tabindex='16' value='" & rs("stream_num") & "'></td>"
				End If
			End If

			'Dropdown for drawings.
			If session("err") = "dwg_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='dwg_id' tabindex='17'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='dwg_id' tabindex='17'>"
			End If
			sqlString = "SELECT dwg_id,dwg_name FROM drawings WHERE dwg_name is not null and dwg_type_id=1 ORDER BY dwg_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("dwg_id") <> "" Then
						If CLng(Request("dwg_id")) = rs2("dwg_id") Then
							response.write "<option value='" & rs2("dwg_id") & "' selected>" & rs2("dwg_name")
						Else
							response.write "<option value='" & rs2("dwg_id") & "'>" & rs2("dwg_name")
						End If
					Else
						If rs("dwg") = rs2("dwg_name") Then
							response.write "<option value='" & rs2("dwg_id") & "' selected>" & rs2("dwg_name")
						Else
							response.write "<option value='" & rs2("dwg_id") & "'>" & rs2("dwg_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			'Dropdown for pipe class.
			If session("err") = "pipe_class_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='pipe_class_id' tabindex='18'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='pipe_class_id' tabindex='18'>"
			End If
			sqlString = "SELECT pipe_class_id,pipe_class_name FROM pipe_classes ORDER BY pipe_class_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("pipe_class_id") <> "" Then
						If CLng(Request("pipe_class_id")) = rs2("pipe_class_id") Then
							response.write "<option value='" & rs2("pipe_class_id") & "' selected>" & rs2("pipe_class_name")
						Else
							response.write "<option value='" & rs2("pipe_class_id") & "'>" & rs2("pipe_class_name")
						End If
					Else
						If rs("pipe_class") = rs2("pipe_class_name") Then
							response.write "<option value='" & rs2("pipe_class_id") & "' selected>" & rs2("pipe_class_name")
						Else
							response.write "<option value='" & rs2("pipe_class_id") & "'>" & rs2("pipe_class_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' tabindex='19' title='Apply changes to this record'></td>"
			End If

			If access = "delete" Then
				recid = rs("line_id")
				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' tabindex='20' title='Delete this record'>Delete</a></td>"
			End If
		Else
			'Draw the second history row.
			If Trim(rs("pipe_std_id")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("pipe_std_id") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			
			If Trim(rs("ansi_din")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("ansi_din") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			
			If Trim(rs("pipe_size")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("pipe_size") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			
			If rs("pipe_mat") <> "" Then
				response.write "<td id='mediumtd'>" & rs("pipe_mat") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			
			If Trim(rs("stream_num")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("stream_num") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			
			If rs("dwg") <> "" Then
				response.write "<td id='mediumtd'>" & rs("dwg") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			
			If rs("pipe_class") <> "" Then
				response.write "<td id='mediumtd'>" & rs("pipe_class") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			
			If access = "write" Or access = "delete" Then
'				If request("firstnum") <> "" Then
'					response.write "<td id='mediumtd'><a href='lines.asp?firstnum=" & request("firstnum") & "&limit=" & limitNum & "&record_id=" & rs("line_id") & "' title='Edit this record'>Edit</a></td>"
				If request("lastnum") <> "" Then
					response.write "<td id='mediumtd'><a href='lines.asp?lastnum=" & request("lastnum") & "&limit=" & limitNum & "&record_id=" & rs("line_id") & "&sort=" & sortkey & "&direction=" & sortdir & "' tabindex='19' title='Edit this record'>Edit</a></td>"
				Else
					response.write "<td id='mediumtd'><a href='lines.asp?limit=" & limitNum & "&record_id=" & rs("line_id") & "&sort=" & sortkey & "&direction=" & sortdir & "' tabindex='19' title='Edit this record'>Edit</a></td>"
				End If
			End If
			If access = "delete" Then
				recid = rs("line_id")
				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' tabindex='19' title='Delete this record'>Delete</a></td>"
			ElseIf recordid < 0 Then
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
		End If
		response.Write "</tr>"
		Response.Write "</table>"

		Response.Write "<table style='width:100%'>"
		Response.Write "<tr><td style='background-color:blue;font-size:1px;padding-top:0px;padding-bottom:0px;width:100%'>&nbsp;</td></tr>"
		Response.Write "</table>"
		rs.Movenext
	Loop
	rs.Close

	'Display "Prev Page" and "Next Page" links.
'	Response.Write "<table style='width:100%'>"
'	Response.Write "<tr>"
'	If IsNumeric(limitNum) Then
'		If CLng(firstNum) < CLng(maxNum) Then
'			response.write "<td ID='headertd' style='width:45%;text-align:right'><a href='lines.asp?firstnum=" & firstNum & "&limit=" & limitNum & "' title='Open the previous page of records'>&lt;-Prev Page</a></td>"
'		Else
'			Response.Write "<td id='headertd' style='width:45%'>&nbsp;</td>"
'		End If
'	Else
'		Response.Write "<td id='headertd' style='width:45%'>&nbsp;</td>"
'	End If
'	Response.Write "<td id='headertd' style='width:10%'>&nbsp;</td>"
'	If IsNumeric(limitNum) Then
'		If CLng(lastNum) > CLng(minNum) Then
'			response.write "<td ID='headertd' style='width:45%;text-align:left'><a href='lines.asp?lastnum=" & lastNum & "&limit=" & limitNum & "' title='Open the next page of records'>Next Page-&gt;</a></td>"
'		Else
'			Response.Write "<td id='headertd' style='width:45%'>&nbsp;</td>"
'		End If
'	Else
'		Response.Write "<td id='headertd' style='width:45%'>&nbsp;</td>"
'	End If
'	Response.Write "</tr>"
'	Response.Write "</table>"
	
	Response.Write "</form>"
	Response.Write "</body>"

	Set rs = Nothing
	Set rs2 = Nothing
	cn.Close
	Set cn = Nothing

Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If

'session("err") = "NONE"

%>
</html>
