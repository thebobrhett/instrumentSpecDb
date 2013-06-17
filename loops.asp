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
 document.form1.action="loops.asp";
 document.form1.submit()
}
function findRecord(id) {
 var newid = id
 document.form1.action="loops.asp?lastnum="+newid;
 document.form1.submit()
}
</script>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Administer Loops</title>
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
' Keith Brooks - Wednesday, April 28, 2010
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
access = UserAccess("equipment", "loops", currentuser)
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
		sortkey = "loop_id"
	End If
	If request("direction") <> "" Then
		sortdir = request("direction")
	Else
		sortdir = "ASC"
	End If
	If Request("limit") <> "" Then
		limitnum = Request("limit")
	Else
		limitnum = "10"
	End If

	response.write "<table ID='headertable' width='100%'>"
	response.write "<tr>"
	response.write "<td ID='headertd' style='width:40%;text-align:left;vertical-align:top'><a href='adminmenu.asp' title='Open the administration main menu'>Menu</a></td>"
	response.write "<td ID='headertd' style='width:20%;text-align:center;vertical-align:center'><h1/>Edit Loops</td>"
	response.write "<td ID='headertd' style='width:40%;text-align:right;vertical-align:top'><a href='#' onClick='openhelp();return false;' title='Open the Admin Guide'>Help</a></td>"
	response.write "</tr>"

	response.Write "<form action='adminaction.asp' id='form1' method='post' name='form1'>"

	response.write "<tr>"
	'Display a dropdown list for the user to select a plant area.  If he selects
	'one, repost the page and display a list of loops for that area that the
	'user can select to "find".
	Response.Write "<td id='headertd'>"
	Response.Write "<table id='headertd'>"
	Response.Write "<tr>"
	Response.Write "<td id='headertd'>Find Area: "
	Response.Write "<select name='find_area' id='find_area' tabindex='1' onchange='reloadPage();'>"
	sqlString = "SELECT plant_area_id,plant_area_name FROM plant_areas ORDER BY plant_area_name"
	Set rs = cn.Execute(sqlString)
	If Not rs.BOF Then
		rs.MoveFirst
		Response.Write "<option value=''> "
		Do While Not rs.EOF
			If IsNumeric(Request("find_area")) Then
				If CLng(Request("find_area")) = rs("plant_area_id") Then
					response.write "<option value='" & rs("plant_area_id") & "' selected>" & rs("plant_area_name")
				Else
					response.write "<option value='" & rs("plant_area_id") & "'>" & rs("plant_area_name")
				End If
			Else
				response.write "<option value='" & rs("plant_area_id") & "'>" & rs("plant_area_name")
			End If
			rs.MoveNext
		Loop
	End If
	rs.Close
	response.write "</select></td>"
	Response.Write "</tr><tr>"
	
	If Request("find_area") <> "" Then
		response.write "<td id='headertd' style='text-align:left'>Find Loop: "
		Response.Write "<select name='find' id='find' tabindex='2' onchange='findRecord(this.value);'>"
		sqlString = "SELECT loop_id,loop_name FROM loops WHERE plant_area_id=" & Request("find_area") & " ORDER BY loop_name"
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If IsNumeric(Request("find")) Then
					If CLng(Request("find")) = rs("loop_id") Then
					Response.Write "<option value='" & rs("loop_id") & "' selected>" & rs("loop_name")
					Else
						Response.Write "<option value='" & rs("loop_id") & "'>" & rs("loop_name")
					End If
				Else
					Response.Write "<option value='" & rs("loop_id") & "'>" & rs("loop_name")
				End If
				rs.MoveNext
			Loop
		End If
		rs.Close
	Else
		Response.Write "<td id='headertd'>&nbsp;</td>"
	End If
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td>"
	If access = "write" Or access = "delete" Then
		response.write "<td ID='headertd' style='text-align:center'><a href='addloops.asp' tabindex='3' title='Add a new loop record'>Add new record</a></td>"
	Else
		Response.Write "<td id='headertd'>&nbsp;</td>"
	End If
	response.write "<td id='headertd'>Records to display:"
	Response.Write "<select name='limit' id='limit' tabindex='4' onchange='reloadPage();'>"
	If limitnum = "All" Then
		Response.Write "<option value='All' selected>All"
	Else
		Response.Write "<option value='All'>All"
	End If
	If limitnum = "10" Then
		Response.Write "<option value='10' selected>10"
	Else
		Response.Write "<option value='10'>10"
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
'		'Get the max record number.
'		sqlString = "SELECT MAX(loop_id) FROM loops"
'		Set rs = cn.Execute(sqlString)
'		If Not rs.BOF Then
'			rs.MoveFirst
'			maxNum = rs(0)
'		Else
'			maxNum = ""
'		End If
'		rs.Close
'
'		'Get the min record number.
'		sqlString = "SELECT MIN(loop_id) FROM loops"
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
'			sqlString = "SELECT loop_id FROM loops WHERE loop_id > " & Request("firstnum") & " ORDER BY loop_id ASC LIMIT " & limitNum
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
	sqlString = "SELECT loops.*,plant_area_name AS plant_area," & _
			"CONCAT(loop_proc_name,' - ',loop_proc_desc) AS loop_proc," & _
			"loop_type_name AS loop_type,loop_func_name AS loop_func," & _
			"dwg_name AS dwg,equip_name AS equip " & _
			"FROM (((((loops LEFT JOIN plant_areas " & _
			"ON loops.plant_area_id=plant_areas.plant_area_id) " & _
			"LEFT JOIN loop_processes ON loops.loop_proc_id=loop_processes.loop_proc_id) " & _
			"LEFT JOIN loop_types ON loops.loop_type_id=loop_types.loop_type_id) " & _
			"LEFT JOIN loop_functions ON loops.loop_func_id=loop_functions.loop_func_id) " & _
			"LEFT JOIN drawings ON loops.dwg_id=drawings.dwg_id) " & _
			"LEFT JOIN equipment ON loops.loop_equip_id=equipment.equip_id "
	If IsNumeric(limitNum) Then
		If request("lastnum") <> "" Then
			sqlString = sqlString & "WHERE loop_id = " & Request("lastnum") & _
					" ORDER BY " & sortkey & " " & sortdir & " LIMIT " & limitNum
'		ElseIf Request("firstnum") <> "" Then
''			sqlString = sqlString & "WHERE loop_id < " & CStr(CLng(request("firstnum")) + CInt(limitNum) + 1) & _
'			sqlString = sqlString & "WHERE loop_id <= " & firstNum & _
'					" ORDER BY " & sortkey & " " & sortdir & " LIMIT " & limitNum
		Else
			sqlString = sqlString & " ORDER BY " & sortkey & " " & sortdir & " LIMIT " & limitNum
		End If
	Else
		sqlString = sqlString & " ORDER BY " & sortkey & " " & sortdir
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

			response.Write "<th id='mediumth' style='width:8%'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_id&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Loop ID&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_id&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth' style='width:10%'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=plant_area&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Plant Area&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=plant_area&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth' style='width:18%'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_name&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Loop Name&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_name&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth' style='width:8%'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_num&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Loop<br />Number&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_num&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth' style='width:37%'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_proc&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Loop Process&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_proc&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth' style='width:10%'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_type&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Loop Type&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_type&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth' style='width:9%'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_func&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Loop Function&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_func&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "</tr>"

			Response.Write "<tr>"
			response.write "<td id='mediumtd'>&nbsp;</td>"

			'Dropdown for plant area.
			If session("err") = "plant_area_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='plant_area_id' tabindex='5'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='plant_area_id' tabindex='5'>"
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

			If session("err") = "loop_name" Then
				If session("loop_name") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='loop_name' size='20' tabindex='6' value='" & session("loop_name") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='loop_name' size='20' tabindex='6' value=''></td>"
				End If
			Else
				If session("loop_name") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='loop_name' size='20' tabindex='6' value='" & session("loop_name") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='loop_name' size='20' tabindex='6' value=''></td>"
				End If
			End If

			If session("err") = "loop_num" Then
				If session("loop_num") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='loop_num' size='5' tabindex='7' value='" & session("loop_num") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='loop_num' size='5' tabindex='7' value=''></td>"
				End If
			Else
				If session("loop_num") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='loop_num' size='5' tabindex='7' value='" & session("loop_num") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='loop_num' size='5' tabindex='7' value=''></td>"
				End If
			End If

			'Dropdown for loop process.
			If session("err") = "loop_proc_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='loop_proc_id' tabindex='8'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='loop_proc_id' tabindex='8'>"
			End If
			sqlString = "SELECT loop_proc_id,CONCAT(loop_proc_name,' - ',loop_proc_desc) AS loop_proc FROM loop_processes ORDER BY loop_proc"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("loop_proc_id") <> "" Then
						If CLng(Request("loop_proc_id")) = rs2("loop_proc_id") Then
							response.write "<option value='" & rs2("loop_proc_id") & "' selected>" & rs2("loop_proc")
						Else
							response.write "<option value='" & rs2("loop_proc_id") & "'>" & rs2("loop_proc")
						End If
					Else
						If Session("loop_proc_id") <> "" Then
							If CLng(session("loop_proc_id")) = rs2("loop_proc_id") Then
								response.write "<option value='" & rs2("loop_proc_id") & "' selected>" & rs2("loop_proc")
							Else
								response.write "<option value='" & rs2("loop_proc_id") & "'>" & rs2("loop_proc")
							End If
						Else
							response.write "<option value='" & rs2("loop_proc_id") & "'>" & rs2("loop_proc")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			'Dropdown for loop type.
			If session("err") = "loop_type_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='loop_type_id' tabindex='9'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='loop_type_id' tabindex='9'>"
			End If
			sqlString = "SELECT loop_type_id,loop_type_name FROM loop_types ORDER BY loop_type_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("loop_type_id") <> "" Then
						If CLng(Request("loop_type_id")) = rs2("loop_type_id") Then
							response.write "<option value='" & rs2("loop_type_id") & "' selected>" & rs2("loop_type_name")
						Else
							response.write "<option value='" & rs2("loop_type_id") & "'>" & rs2("loop_type_name")
						End If
					Else
						If Session("loop_type_id") <> "" Then
							If CLng(session("loop_type_id")) = rs2("loop_type_id") Then
								response.write "<option value='" & rs2("loop_type_id") & "' selected>" & rs2("loop_type_name")
							Else
								response.write "<option value='" & rs2("loop_type_id") & "'>" & rs2("loop_type_name")
							End If
						Else
							response.write "<option value='" & rs2("loop_type_id") & "'>" & rs2("loop_type_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			'Dropdown for loop function.
			If session("err") = "loop_func_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='loop_func_id' tabindex='10'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='loop_func_id' tabindex='10'>"
			End If
			sqlString = "SELECT loop_func_id,loop_func_name FROM loop_functions ORDER BY loop_func_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("loop_func_id") <> "" Then
						If CLng(Request("loop_func_id")) = rs2("loop_func_id") Then
							response.write "<option value='" & rs2("loop_func_id") & "' selected>" & rs2("loop_func_name")
						Else
							response.write "<option value='" & rs2("loop_func_id") & "'>" & rs2("loop_func_name")
						End If
					Else
						If Session("loop_func_id") <> "" Then
							If CLng(session("loop_func_id")) = rs2("loop_func_id") Then
								response.write "<option value='" & rs2("loop_func_id") & "' selected>" & rs2("loop_func_name")
							Else
								response.write "<option value='" & rs2("loop_func_id") & "'>" & rs2("loop_func_name")
							End If
						Else
							response.write "<option value='" & rs2("loop_func_id") & "'>" & rs2("loop_func_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			'Draw the second header row.
			response.Write "<table width='100%'>"
			response.Write "<tr>"
			response.Write "<th id='mediumth' style='width:6%'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_suff&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Loop<br />Suffix&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_suff&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth' style='width:25%'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_desc&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Loop Description&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_desc&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth' style='width:25%'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_note&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Loop Note&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_note&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth' style='width:14%'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=dwg&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Drawing&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=dwg&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			response.Write "<th id='mediumth' style='width:17%'>"
			Response.Write "<table><tr>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=equip&direction=DESC&limit=" & limitNum & "'>"
			response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
			Response.Write "<th style='font-size:8pt'>&nbsp;Equipment&nbsp;</th>"
			Response.Write "<td id='headertd'><a href='loops.asp?sort=equip&direction=ASC&limit=" & limitNum & "'>"
			Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
			If access = "write" Or access = "delete" Then
 			  response.Write "<th id='mediumth' style='width:8%'>&nbsp;</th>"
 			End If
			If access = "delete" Or recordid < 0 Then
 			  response.Write "<th id='mediumth' style='width:5%'>&nbsp;</th>"
 			End If
			response.Write "</tr>"

			If session("err") = "loop_suff" Then
				If session("loop_suff") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='loop_suff' size='5' tabindex='11' value='" & session("loop_suff") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='loop_suff' size='5' tabindex='11' value=''></td>"
				End If
			Else
				If session("loop_suff") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='loop_suff' size='5' tabindex='11' value='" & session("loop_suff") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='loop_suff' size='5' tabindex='11' value=''></td>"
				End If
			End If

			If session("err") = "loop_desc" Then
				If session("loop_desc") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='loop_desc' rows='2' cols='30' tabindex='12'>" & session("loop_desc") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='loop_desc' rows='2' cols='30' tabindex='12'></textarea></td>"
				End If
			Else
				If session("loop_desc") <> "" Then
					response.write "<td id='mediumtd'><textarea name='loop_desc' rows='2' cols='30' tabindex='12'>" & session("loop_desc") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd'><textarea name='loop_desc' rows='2' cols='30' tabindex='12'></textarea></td>"
				End If
			End If

			If session("err") = "loop_note" Then
				If session("loop_note") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='loop_note' rows='2' cols='30' tabindex='13'>" & session("loop_note") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='loop_note' rows='2' cols='30' tabindex='13'></textarea></td>"
				End If
			Else
				If session("loop_note") <> "" Then
					response.write "<td id='mediumtd'><textarea name='loop_note' rows='2' cols='30' tabindex='13'>" & session("loop_note") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd'><textarea name='loop_note' rows='2' cols='30' tabindex='13'></textarea></td>"
				End If
			End If

			'Dropdown for drawings.
			If session("err") = "dwg_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='dwg_id' tabindex='14'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='dwg_id' tabindex='14'>"
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

			'Dropdown for equipment.
			If session("err") = "loop_equip_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='loop_equip_id' tabindex='15'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='loop_equip_id' tabindex='15'>"
			End If
			sqlString = "SELECT equip_id,equip_name FROM equipment ORDER BY equip_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("loop_equip_id") <> "" Then
						If CLng(Request("loop_equip_id")) = rs2("equip_id") Then
							response.write "<option value='" & rs2("equip_id") & "' selected>" & rs2("equip_name")
						Else
							response.write "<option value='" & rs2("equip_id") & "'>" & rs2("equip_name")
						End If
					Else
						If Session("loop_equip_id") <> "" Then
							If CLng(session("loop_equip_id")) = rs2("equip_id") Then
								response.write "<option value='" & rs2("equip_id") & "' selected>" & rs2("equip_name")
							Else
								response.write "<option value='" & rs2("equip_id") & "'>" & rs2("equip_name")
							End If
						Else
							response.write "<option value='" & rs2("equip_id") & "'>" & rs2("equip_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' tabindex='16' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"

			Response.Write "<td id='mediumtd' style='text-align:center'><a href='loops.asp?limit=" & limitnum & "' tabindex='17' title='Cancel changes to this record'>Cancel</a></td>"

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
			firstNum = rs("loop_id")
		End If
		firstPass = False
		lastNum = rs("loop_id")

		'Draw the first header row.
		response.Write "<table width='100%'>"
		response.Write "<tr>"
		response.Write "<th id='mediumth' style='width:8%'>"
		Response.Write "<table><tr>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_id&direction=DESC&limit=" & limitNum & "'>"
		response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
		Response.Write "<th style='font-size:8pt'>&nbsp;Loop ID&nbsp;</th>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_id&direction=ASC&limit=" & limitNum & "'>"
		Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		response.Write "<th id='mediumth' style='width:10%'>"
		Response.Write "<table><tr>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=plant_area&direction=DESC&limit=" & limitNum & "'>"
		response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
		Response.Write "<th style='font-size:8pt'>&nbsp;Plant Area&nbsp;</th>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=plant_area&direction=ASC&limit=" & limitNum & "'>"
		Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		response.Write "<th id='mediumth' style='width:18%'>"
		Response.Write "<table><tr>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_name&direction=DESC&limit=" & limitNum & "'>"
		response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
		Response.Write "<th style='font-size:8pt'>&nbsp;Loop Name&nbsp;</th>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_name&direction=ASC&limit=" & limitNum & "'>"
		Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		response.Write "<th id='mediumth' style='width:8%'>"
		Response.Write "<table><tr>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_num&direction=DESC&limit=" & limitNum & "'>"
		response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
		Response.Write "<th style='font-size:8pt'>&nbsp;Loop<br />Number&nbsp;</th>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_num&direction=ASC&limit=" & limitNum & "'>"
		Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		response.Write "<th id='mediumth' style='width:37%'>"
		Response.Write "<table><tr>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_proc&direction=DESC&limit=" & limitNum & "'>"
		response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
		Response.Write "<th style='font-size:8pt'>&nbsp;Loop Process&nbsp;</th>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_proc&direction=ASC&limit=" & limitNum & "'>"
		Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		response.Write "<th id='mediumth' style='width:10%'>"
		Response.Write "<table><tr>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_type&direction=DESC&limit=" & limitNum & "'>"
		response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
		Response.Write "<th style='font-size:8pt'>&nbsp;Loop Type&nbsp;</th>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_type&direction=ASC&limit=" & limitNum & "'>"
		Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		response.Write "<th id='mediumth' style='width:9%'>"
		Response.Write "<table><tr>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_func&direction=DESC&limit=" & limitNum & "'>"
		response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
		Response.Write "<th style='font-size:8pt'>&nbsp;Loop Function&nbsp;</th>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_func&direction=ASC&limit=" & limitNum & "'>"
		Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		response.Write "</tr>"

		Response.Write "<tr>"
		If CLng(rs("loop_id")) = CLng(recordid) Then
			'Draw the first data entry line
			response.write "<td id='mediumtd'>" & rs("loop_id") & "</td>"

			If rs("plant_area") <> "" Then
				response.write "<td id='mediumtd'>" & rs("plant_area") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("loop_name")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("loop_name") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("loop_num")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("loop_num") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("loop_proc") <> "" Then
				response.write "<td id='mediumtd'>" & rs("loop_proc") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If

			'Dropdown for plant area.
'			If session("err") = "plant_area_id" Then
'				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='plant_area_id'>"
'			Else
'				response.write "<td id='mediumtd' style='text-align:center'><select name='plant_area_id'>"
'			End If
'			sqlString = "SELECT plant_area_id,plant_area_name FROM plant_areas ORDER BY plant_area_name"
'			Set rs2 = cn.Execute(sqlString)
'			If Not rs2.BOF Then
'				rs2.MoveFirst
'				Response.Write "<option value=''> "
'				Do While Not rs2.EOF
'					If Request("plant_area_id") <> "" Then
'						If CLng(Request("plant_area_id")) = rs2("plant_area_id") Then
'							response.write "<option value='" & rs2("plant_area_id") & "' selected>" & rs2("plant_area_name")
'						Else
'							response.write "<option value='" & rs2("plant_area_id") & "'>" & rs2("plant_area_name")
'						End If
'					Else
'						If rs("plant_area") = rs2("plant_area_name") Then
'							response.write "<option value='" & rs2("plant_area_id") & "' selected>" & rs2("plant_area_name")
'						Else
'							response.write "<option value='" & rs2("plant_area_id") & "'>" & rs2("plant_area_name")
'						End If
'					End If
'					rs2.MoveNext
'				Loop
'			End If
'			rs2.Close
'			response.write "</select></td>"
'
'			If session("err") = "loop_name" Then
'				If session("loop_name") <> "" Then
'					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='loop_name' size='20' value='" & session("loop_name") & "'></td>"
'				Else
'					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='loop_name' size='20' value='" & rs("loop_name") & "'></td>"
'				End If
'			Else
'				If session("loop_name") <> "" Then
'					response.write "<td id='mediumtd'><input type='text' name='loop_name' size='20' value='" & session("loop_name") & "'></td>"
'				Else
'					response.write "<td id='mediumtd'><input type='text' name='loop_name' size='20' value='" & rs("loop_name") & "'></td>"
'				End If
'			End If
'
'			If session("err") = "loop_num" Then
'				If session("loop_num") <> "" Then
'					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='loop_num' size='5' value='" & session("loop_num") & "'></td>"
'				Else
'					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='loop_num' size='5' value='" & rs("loop_num") & "'></td>"
'				End If
'			Else
'				If session("loop_num") <> "" Then
'					response.write "<td id='mediumtd'><input type='text' name='loop_num' size='5' value='" & session("loop_num") & "'></td>"
'				Else
'					response.write "<td id='mediumtd'><input type='text' name='loop_num' size='5' value='" & rs("loop_num") & "'></td>"
'				End If
'			End If
'
'			'Dropdown for loop process.
'			If session("err") = "loop_proc_id" Then
'				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='loop_proc_id'>"
'			Else
'				response.write "<td id='mediumtd' style='text-align:center'><select name='loop_proc_id'>"
'			End If
'			sqlString = "SELECT loop_proc_id,CONCAT(loop_proc_name,' - ',loop_proc_desc) AS loop_proc FROM loop_processes ORDER BY loop_proc"
'			Set rs2 = cn.Execute(sqlString)
'			If Not rs2.BOF Then
'				rs2.MoveFirst
'				Response.Write "<option value=''> "
'				Do While Not rs2.EOF
'					If Request("loop_proc_id") <> "" Then
'						If CLng(Request("loop_proc_id")) = rs2("loop_proc_id") Then
'							response.write "<option value='" & rs2("loop_proc_id") & "' selected>" & rs2("loop_proc")
'						Else
'							response.write "<option value='" & rs2("loop_proc_id") & "'>" & rs2("loop_proc")
'						End If
'					Else
'						If rs("loop_proc") = rs2("loop_proc") Then
'							response.write "<option value='" & rs2("loop_proc_id") & "' selected>" & rs2("loop_proc")
'						Else
'							response.write "<option value='" & rs2("loop_proc_id") & "'>" & rs2("loop_proc")
'						End If
'					End If
'					rs2.MoveNext
'				Loop
'			End If
'			rs2.Close
'			response.write "</select></td>"

			'Dropdown for loop type.
			If session("err") = "loop_type_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='loop_type_id' tabindex='5'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='loop_type_id' tabindex='5'>"
			End If
			sqlString = "SELECT loop_type_id,loop_type_name FROM loop_types ORDER BY loop_type_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("loop_type_id") <> "" Then
						If CLng(Request("loop_type_id")) = rs2("loop_type_id") Then
							response.write "<option value='" & rs2("loop_type_id") & "' selected>" & rs2("loop_type_name")
						Else
							response.write "<option value='" & rs2("loop_type_id") & "'>" & rs2("loop_type_name")
						End If
					Else
						If rs("loop_type") = rs2("loop_type_name") Then
							response.write "<option value='" & rs2("loop_type_id") & "' selected>" & rs2("loop_type_name")
						Else
							response.write "<option value='" & rs2("loop_type_id") & "'>" & rs2("loop_type_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			'Dropdown for loop function.
			If session("err") = "loop_func_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='loop_func_id' tabindex='6'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='loop_func_id' tabindex='6'>"
			End If
			sqlString = "SELECT loop_func_id,loop_func_name FROM loop_functions ORDER BY loop_func_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("loop_func_id") <> "" Then
						If CLng(Request("loop_func_id")) = rs2("loop_func_id") Then
							response.write "<option value='" & rs2("loop_func_id") & "' selected>" & rs2("loop_func_name")
						Else
							response.write "<option value='" & rs2("loop_func_id") & "'>" & rs2("loop_func_name")
						End If
					Else
						If rs("loop_func") = rs2("loop_func_name") Then
							response.write "<option value='" & rs2("loop_func_id") & "' selected>" & rs2("loop_func_name")
						Else
							response.write "<option value='" & rs2("loop_func_id") & "'>" & rs2("loop_func_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

		Else
			'Draw the first history row.
			response.write "<td id='mediumtd'>" & rs("loop_id") & "</td>"
			If rs("plant_area") <> "" Then
				response.write "<td id='mediumtd'>" & rs("plant_area") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("loop_name")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("loop_name") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("loop_num")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("loop_num") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("loop_proc") <> "" Then
				response.write "<td id='mediumtd'>" & rs("loop_proc") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("loop_type") <> "" Then
				response.write "<td id='mediumtd'>" & rs("loop_type") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("loop_func") <> "" Then
				response.write "<td id='mediumtd'>" & rs("loop_func") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
		End If
		response.write "</tr>"
		Response.Write "</table>"

		'Draw the second header row.
		response.Write "<table width='100%'>"
		response.Write "<tr>"
		response.Write "<th id='mediumth' style='width:6%'>"
		Response.Write "<table><tr>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_suff&direction=DESC&limit=" & limitNum & "'>"
		response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
		Response.Write "<th style='font-size:8pt'>&nbsp;Loop<br />Suffix&nbsp;</th>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_suff&direction=ASC&limit=" & limitNum & "'>"
		Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		response.Write "<th id='mediumth' style='width:25%'>"
		Response.Write "<table><tr>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_desc&direction=DESC&limit=" & limitNum & "'>"
		response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
		Response.Write "<th style='font-size:8pt'>&nbsp;Loop Description&nbsp;</th>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_desc&direction=ASC&limit=" & limitNum & "'>"
		Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		response.Write "<th id='mediumth' style='width:25%'>"
		Response.Write "<table><tr>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_note&direction=DESC&limit=" & limitNum & "'>"
		response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
		Response.Write "<th style='font-size:8pt'>&nbsp;Loop Note&nbsp;</th>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=loop_note&direction=ASC&limit=" & limitNum & "'>"
		Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		response.Write "<th id='mediumth' style='width:14%'>"
		Response.Write "<table><tr>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=dwg&direction=DESC&limit=" & limitNum & "'>"
		response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
		Response.Write "<th style='font-size:8pt'>&nbsp;Drawing&nbsp;</th>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=dwg&direction=ASC&limit=" & limitNum & "'>"
		Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		response.Write "<th id='mediumth' style='width:17%'>"
		Response.Write "<table><tr>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=equip&direction=DESC&limit=" & limitNum & "'>"
		response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
		Response.Write "<th style='font-size:8pt'>&nbsp;Equipment&nbsp;</th>"
		Response.Write "<td id='headertd'><a href='loops.asp?sort=equip&direction=ASC&limit=" & limitNum & "'>"
		Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
		If access = "write" Or access = "delete" Then
 		  response.Write "<th id='mediumth' style='width:8%'>&nbsp;</th>"
 		End If
		If access = "delete" Or recordid < 0 Then
 		  response.Write "<th id='mediumth' style='width:5%'>&nbsp;</th>"
 		End If
		response.Write "</tr>"

		Response.Write "<tr>"
		If CLng(rs("loop_id")) = CLng(recordid) Then
			'Draw the second data entry line
			If session("err") = "loop_suff" Then
				If session("loop_suff") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='loop_suff' size='5' tabindex='7' value='" & session("loop_suff") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='loop_suff' size='5' tabindex='7' value='" & rs("loop_suff") & "'></td>"
				End If
			Else
				If session("loop_suff") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='loop_suff' size='5' tabindex='7' value='" & session("loop_suff") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='loop_suff' size='5' tabindex='7' value='" & rs("loop_suff") & "'></td>"
				End If
			End If

			If session("err") = "loop_desc" Then
				If session("loop_desc") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='loop_desc' rows='2' cols='35' tabindex='8'>" & session("loop_desc") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='loop_desc' rows='2' cols='35' tabindex='8'>" & rs("loop_desc") & "</textarea></td>"
				End If
			Else
				If session("loop_desc") <> "" Then
					response.write "<td id='mediumtd'><textarea name='loop_desc' rows='2' cols='35' tabindex='8'>" & session("loop_desc") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd'><textarea name='loop_desc' rows='2' cols='35' tabindex='8'>" & rs("loop_desc") & "</textarea></td>"
				End If
			End If

			If session("err") = "loop_note" Then
				If session("loop_note") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='loop_note' rows='2' cols='35' tabindex='9'>" & session("loop_note") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='loop_note' rows='2' cols='35' tabindex='9'>" & rs("loop_note") & "</textarea></td>"
				End If
			Else
				If session("loop_note") <> "" Then
					response.write "<td id='mediumtd'><textarea name='loop_note' rows='2' cols='35' tabindex='9'>" & session("loop_note") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd'><textarea name='loop_note' rows='2' cols='35' tabindex='9'>" & rs("loop_note") & "</textarea></td>"
				End If
			End If

			'Dropdown for drawings.
			If session("err") = "dwg_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='dwg_id' tabindex='10'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='dwg_id' tabindex='10'>"
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

			If rs("equip") <> "" Then
				response.write "<td id='mediumtd'>" & rs("equip") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			'Dropdown for equipment.
'			If session("err") = "loop_equip_id" Then
'				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='loop_equip_id'>"
'			Else
'				response.write "<td id='mediumtd' style='text-align:center'><select name='loop_equip_id'>"
'			End If
'			sqlString = "SELECT equip_id,equip_name FROM equipment ORDER BY equip_name"
'			Set rs2 = cn.Execute(sqlString)
'			If Not rs2.BOF Then
'				rs2.MoveFirst
'				Response.Write "<option value=''> "
'				Do While Not rs2.EOF
'					If Request("loop_equip_id") <> "" Then
'						If CLng(Request("loop_equip_id")) = rs2("equip_id") Then
'							response.write "<option value='" & rs2("equip_id") & "' selected>" & rs2("equip_name")
'						Else
'							response.write "<option value='" & rs2("equip_id") & "'>" & rs2("equip_name")
'						End If
'					Else
'						If rs("equip") = rs2("equip_name") Then
'							response.write "<option value='" & rs2("equip_id") & "' selected>" & rs2("equip_name")
'						Else
'							response.write "<option value='" & rs2("equip_id") & "'>" & rs2("equip_name")
'						End If
'					End If
'					rs2.MoveNext
'				Loop
'			End If
'			rs2.Close
'			response.write "</select></td>"

			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' tabindex='11' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"
			End If

			If access = "delete" Then
				recid = rs("loop_id")
				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' tabindex='11' title='Delete this record'>Delete</a></td>"
			End If
		Else
			'Draw the second history row.
			If Trim(rs("loop_suff")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("loop_suff") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("loop_desc")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("loop_desc") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			
			If Trim(rs("loop_note")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("loop_note") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			
			If rs("dwg") <> "" Then
				response.write "<td id='mediumtd'>" & rs("dwg") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			
			If rs("equip") <> "" Then
				response.write "<td id='mediumtd'>" & rs("equip") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			
			If access = "write" Or access = "delete" Then
'				If request("firstnum") <> "" Then
'					response.write "<td id='mediumtd'><a href='loops.asp?firstnum=" & request("firstnum") & "&limit=" & limitNum & "&record_id=" & rs("loop_id") & "' title='Edit this record'>Edit</a></td>"
				If request("lastnum") <> "" Then
					response.write "<td id='mediumtd'><a href='loops.asp?lastnum=" & request("lastnum") & "&limit=" & limitNum & "&record_id=" & rs("loop_id") & "&sort=" & sortkey & "&direction=" & sortdir & "' tabindex='11' title='Edit this record'>Edit</a></td>"
				Else
					response.write "<td id='mediumtd'><a href='loops.asp?limit=" & limitNum & "&record_id=" & rs("loop_id") & "&sort=" & sortkey & "&direction=" & sortdir & "' tabindex='11' title='Edit this record'>Edit</a></td>"
				End If
			End If
			If access = "delete" Then
				recid = rs("loop_id")
				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' tabindex='11' title='Delete this record'>Delete</a></td>"
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
'			response.write "<td ID='headertd' style='width:45%;text-align:right'><a href='loops.asp?firstnum=" & firstNum & "&limit=" & limitNum & "' title='Open the previous page of records'>&lt;-Prev Page</a></td>"
'		Else
'			Response.Write "<td id='headertd' style='width:45%'>&nbsp;</td>"
'		End If
'	Else
'		Response.Write "<td id='headertd' style='width:45%'>&nbsp;</td>"
'	End If
'	Response.Write "<td id='headertd' style='width:10%'>&nbsp;</td>"
'	If IsNumeric(limitNum) Then
'		If CLng(lastNum) > CLng(minNum) Then
'			response.write "<td ID='headertd' style='width:45%;text-align:left'><a href='loops.asp?lastnum=" & lastNum & "&limit=" & limitNum & "' title='Open the next page of records'>Next Page-&gt;</a></td>"
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
