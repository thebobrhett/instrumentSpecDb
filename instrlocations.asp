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
 document.form1.action="instrlocations.asp";
 document.form1.submit()
}
</script>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Administer Instrument Locations</title>
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
' Keith Brooks - Monday, April 26, 2010
'   Creation.
'*************

'on error resume next

dim cn
dim rs
Dim rs2
dim recordid
Dim currentuser
Dim access
Dim recid
Dim sqlString
Dim sortkey
Dim sortdir
Dim limitnum

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("equipment","instrlocations",currentuser)
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
		sortkey = "instr_location_id"
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
	response.write "<td ID='headertd' style='width:20%;text-align:left;vertical-align:top'><a href='adminmenu.asp' title='Open the administration main menu'>Menu</a></td>"
	response.write "<td ID='headertd' style='width:60%;text-align:center;vertical-align:center'><h1/>Edit Instrument Location List</td>"
	response.write "<td ID='headertd' style='width:20%;text-align:right;vertical-align:top'><a href='#' onClick='openhelp();return false;' title='Open the Admin Guide'>Help</a></td>"
	response.write "</tr>"

	response.Write "<form action='adminaction.asp' id='form1' method='post' name='form1'>"

	response.write "<tr>"
	response.write "<td id='headertd'>&nbsp;</td>"
	If access = "write" Or access = "delete" Then
		response.write "<td ID='headertd' style='text-align:center'><a href='instrlocations.asp?record_id=-1&sort=" & sortkey & "&direction=" & sortdir & "&limit=" & limitnum & "' title='Add a new instrument location record'>Add new record</a></td>"
	Else
		Response.Write "<td id='headertd'>&nbsp;</td>"
	End If
	response.write "<td id='headertd'>Records to display:"
	Response.Write "<select name='limit' id='limit' onchange='reloadPage();'>"
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
		Response.Write "<option value='50' selected>50"
	Else
		Response.Write "<option value='50'>50"
	End If
	If limitnum = "100" Then
		Response.Write "<option value='100' selected>100"
	Else
		Response.Write "<option value='100'>100"
	End If
	Response.Write "</select></td>"
	response.write "</tr>"
	response.write "</table>"

	Response.Write "<br />"

	'Draw the header
	Response.Write "<div style='text-align:center'>"
	response.Write "<table width='100%'>"
	response.Write "<tr>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='instrlocations.asp?sort=instr_location_id&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'&nbsp;>Location<br />ID&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='instrlocations.asp?sort=instr_location_id&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='instrlocations.asp?sort=plant_area&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Plant<br />Area&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='instrlocations.asp?sort=plant_area&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='instrlocations.asp?sort=instr_location_name&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Location<br />Name&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='instrlocations.asp?sort=instr_location_name&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='instrlocations.asp?sort=instr_location_desc&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Location<br />Description&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='instrlocations.asp?sort=instr_location_desc&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"

	If access = "write" Or access = "delete" Then
 	  response.Write "<th id='mediumth'>&nbsp;</th>"
 	End If
	If access = "delete" Or recordid < 0 Then
 	  response.Write "<th id='mediumth'>&nbsp;</th>"
 	End If
	response.Write "</tr>"

	response.write "<input type='hidden' name='RECORD' value='" & recordid & "'>"
	response.write "<input type='hidden' name='SORT' value='" & sortkey & "'>"
	response.write "<input type='hidden' name='DIRECTION' value='" & sortdir & "'>"

	'Read the form data.
	If limitNum = "All" Then
		sqlString = "SELECT instr_location_id,plant_area_name AS plant_area," & _
				"instr_location_name,instr_location_desc " & _
				"FROM instrument_locations LEFT JOIN plant_areas " & _
				"ON instrument_locations.plant_area_id=plant_areas.plant_area_id " & _
				"ORDER BY " & sortkey & " " & sortdir
	Else
		sqlString = "SELECT instr_location_id,plant_area_name AS plant_area," & _
				"instr_location_name,instr_location_desc " & _
				"FROM instrument_locations LEFT JOIN plant_areas " & _
				"ON instrument_locations.plant_area_id=plant_areas.plant_area_id " & _
				"ORDER BY " & sortkey & " " & sortdir & " LIMIT " & limitnum
	End If

	set rs = cn.Execute(sqlString)

	If Not rs.BOF Then
	  rs.MoveFirst
	End If
	  
	'If recordid<0, the user has selected "Add new record" so insert a blank data entry line
	'at the top of the form.
	If access = "write" Or access = "delete" Then
		If recordid < 0 Then
			Response.Write "<tr>"

			response.write "<td id='mediumtd'>&nbsp;</td>"

			'Dropdown for plant area.
			If session("err") = "plant_area_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='plant_area_id'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='plant_area_id'>"
			End If
			sqlString = "SELECT plant_area_id,plant_area_name FROM plant_areas ORDER BY plant_area_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("plant_area_id") <> "" Then
						If CInt(Request("plant_area_id")) = rs2("plant_area_id") Then
							response.write "<option value='" & rs2("plant_area_id") & "' selected>" & rs2("plant_area_name")
						Else
							response.write "<option value='" & rs2("plant_area_id") & "'>" & rs2("plant_area_name")
						End If
					Else
						If Session("plant_area_id") <> "" Then
							If CInt(session("plant_area_id")) = rs2("plant_area_id") Then
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

			If session("err") = "instr_location_name" Then
				If session("instr_location_name") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='instr_location_name' size='50' value='" & session("instr_location_name") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='instr_location_name' size='50' value=''></td>"
				End If
			Else
				If session("instr_location_name") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='instr_location_name' size='50' value='" & session("instr_location_name") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='instr_location_name' size='50' value=''></td>"
				End If
			End If

			If session("err") = "instr_location_desc" Then
				If session("instr_location_desc") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='instr_location_desc' cols='40' rows='2'>" & session("instr_location_desc") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='instr_location_desc' cols='40' rows='2'></textarea></td>"
				End If
			Else
				If session("instr_location_desc") <> "" Then
					response.write "<td id='mediumtd'><textarea name='instr_location_desc' cols='40' rows='2'>" & session("instr_location_desc") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd'><textarea name='instr_location_desc' cols='40' rows='2'></textarea></td>"
				End If
			End If

			response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"

			Response.Write "<td id='mediumtd' style='text-align:center'><a href='instrlocations.asp?sort=" & sortkey & "&direction=" & sortdir & "&limit=" & limitnum & "' title='Cancel changes to this record'>Cancel</a></td>"

			Response.Write "</tr>"
		End If
	End If

	Do While Not rs.EOF
		Response.Write "<tr>"
		If CLng(rs("instr_location_id")) = CLng(recordid) Then
			'Draw the data entry line
			response.write "<td id='mediumtd'>" & rs("instr_location_id") & "</td>"

			'Dropdown for plant area.
			If session("err") = "plant_area_id" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='plant_area_id'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='plant_area_id'>"
			End If
			sqlString = "SELECT plant_area_id,plant_area_name FROM plant_areas ORDER BY plant_area_name"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("plant_area_id") <> "" Then
						If CInt(Request("plant_area_id")) = rs2("plant_area_id") Then
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

			If session("err") = "instr_location_name" Then
				If session("instr_location_name") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='instr_location_name' size='50' value='" & session("instr_location_name") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='instr_location_name' size='50' value='" & rs("instr_location_name") & "'></td>"
				End If
			Else
				If session("instr_location_name") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='instr_location_name' size='50' value='" & session("instr_location_name") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='instr_location_name' size='50' value='" & rs("instr_location_name") & "'></td>"
				End If
			End If

			If session("err") = "instr_location_desc" Then
				If session("instr_location_desc") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='instr_location_desc' cols='40' rows='2'>" & session("instr_location_desc") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='instr_location_desc' cols='40' rows='2'>" & rs("instr_location_desc") & "</textarea></td>"
				End If
			Else
				If session("instr_location_desc") <> "" Then
					response.write "<td id='mediumtd'><textarea name='instr_location_desc' cols='40' rows='2'>" & session("instr_location_desc") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd'><textarea name='instr_location_desc' cols='40' rows='2'>" & rs("instr_location_desc") & "</textarea></td>"
				End If
			End If

			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"
			End If

			If access = "delete" Then
				recid = rs("instr_location_id")
				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
			End If
		Else
			'Draw the history records
			response.write "<tr>"
			response.write "<td id='mediumtd'>" & rs("instr_location_id") & "</td>"
			If rs("plant_area") <> "" Then
				response.write "<td id='mediumtd'>" & rs("plant_area") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("instr_location_name") <> "" Then
				response.write "<td id='mediumtd'>" & rs("instr_location_name") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If rs("instr_location_desc") <> "" And rs("instr_location_desc") <> " " Then
				response.write "<td id='mediumtd'>" & rs("instr_location_desc") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd'><a href='instrlocations.asp?record_id=" & rs("instr_location_id") & "&sort=" & sortkey & "&direction=" & sortdir & "&limit=" & limitnum & "' title='Edit this record'>Edit</a></td>"
			End If
			If access = "delete" Then
				recid = rs("instr_location_id")
				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
			ElseIf recordid < 0 Then
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			response.write "</tr>"
		End If
		rs.Movenext
	loop
	rs.Close
	
	Response.Write "</form>"
	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "</body>"

	'session("err") = "NONE"

	Set rs = Nothing
	cn.Close
	Set cn = Nothing

Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If
%>
</html>
