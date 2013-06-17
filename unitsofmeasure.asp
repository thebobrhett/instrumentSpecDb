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
 document.form1.action="unitsofmeasure.asp";
 document.form1.submit()
}
</script>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Administer Units of Measure</title>
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
' Keith Brooks - Tuesday, May 4, 2010
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
access = UserAccess("equipment","unitsofmeasure",currentuser)
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
		sortkey = "uom_id"
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
	response.write "<td ID='headertd' style='width:60%;text-align:center;vertical-align:center'><h1/>Edit Units of Measure</td>"
	response.write "<td ID='headertd' style='width:20%;text-align:right;vertical-align:top'><a href='#' onClick='openhelp();return false;' title='Open the Admin Guide'>Help</a></td>"
	response.write "</tr>"

	response.Write "<form action='adminaction.asp' id='form1' method='post' name='form1'>"

	response.write "<tr>"
	response.write "<td id='headertd'>&nbsp;</td>"
	If access = "write" Or access = "delete" Then
		response.write "<td ID='headertd' style='text-align:center'><a href='unitsofmeasure.asp?record_id=-1&sort=" & sortkey & "&direction=" & sortdir & "&limit=" & limitnum & "' title='Add a new instrument function type record'>Add new record</a></td>"
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
	If limitnum = "50" Then
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
	Response.Write "<td id='headertd'><a href='unitsofmeasure.asp?sort=uom_id&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'&nbsp;>Unit of<br />Measure<br />ID&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='unitsofmeasure.asp?sort=uom_id&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='unitsofmeasure.asp?sort=uom_code&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Unit of<br />Measure<br />Code&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='unitsofmeasure.asp?sort=uom_code&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='unitsofmeasure.asp?sort=uom_name&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Unit of<br />Measure<br />Name&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='unitsofmeasure.asp?sort=uom_name&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='unitsofmeasure.asp?sort=uom_desc&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Unit of<br />Measure<br />Description&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='unitsofmeasure.asp?sort=uom_desc&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='unitsofmeasure.asp?sort=uom_type&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Unit of<br />Measure<br />Type&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='unitsofmeasure.asp?sort=uom_type&direction=ASC&limit=" & limitNum & "'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='unitsofmeasure.asp?sort=uom_kind&direction=DESC&limit=" & limitNum & "'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Unit of<br />Measure<br />Kind&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='unitsofmeasure.asp?sort=uom_kind&direction=ASC&limit=" & limitNum & "'>"
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
		sqlString = "SELECT * FROM units_of_measure " & _
				"ORDER BY " & sortkey & " " & sortdir
	Else
		sqlString = "SELECT * FROM units_of_measure " & _
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

			If session("err") = "uom_code" Then
				If session("uom_code") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='uom_code' size='10' value='" & session("uom_code") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='uom_code' size='10' value=''></td>"
				End If
			Else
				If session("uom_code") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='uom_code' size='10' value='" & session("uom_code") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='uom_code' size='10' value=''></td>"
				End If
			End If

			If session("err") = "uom_name" Then
				If session("uom_name") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='uom_name' size='20' value='" & session("uom_name") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='uom_name' size='20' value=''></td>"
				End If
			Else
				If session("uom_name") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='uom_name' size='20' value='" & session("uom_name") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='uom_name' size='20' value=''></td>"
				End If
			End If

			If session("err") = "uom_desc" Then
				If session("uom_desc") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='uom_desc' cols='30' rows='2'>" & session("uom_desc") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='uom_desc' cols='30' rows='2'></textarea></td>"
				End If
			Else
				If session("uom_desc") <> "" Then
					response.write "<td id='mediumtd'><textarea name='uom_desc' cols='30' rows='2'>" & session("uom_desc") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd'><textarea name='uom_desc' cols='30' rows='2'></textarea></td>"
				End If
			End If

			'Dropdown for unit of measure type.
			If session("err") = "uom_type" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='uom_type'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='uom_type'>"
			End If
			sqlString = "SELECT uom_type_id,CONCAT(uom_type_id,' - ',uom_type_desc) AS uom_type_name FROM unit_of_measure_types ORDER BY uom_type_id"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("uom_type") <> "" Then
						If CInt(Request("uom_type")) = rs2("uom_type_id") Then
							response.write "<option value='" & rs2("uom_type_id") & "' selected>" & rs2("uom_type_name")
						Else
							response.write "<option value='" & rs2("uom_type_id") & "'>" & rs2("uom_type_name")
						End If
					Else
						If Session("uom_type") <> "" Then
							If CInt(session("uom_type")) = rs2("uom_type_id") Then
								response.write "<option value='" & rs2("uom_type_id") & "' selected>" & rs2("uom_type_name")
							Else
								response.write "<option value='" & rs2("uom_type_id") & "'>" & rs2("uom_type_name")
							End If
						Else
							response.write "<option value='" & rs2("uom_type_id") & "'>" & rs2("uom_type_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			If session("err") = "uom_kind" Then
				If session("uom_kind") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='uom_kind' size='5' value='" & session("uom_kind") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='uom_kind' size='5' value=''></td>"
				End If
			Else
				If session("uom_kind") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='uom_kind' size='5' value='" & session("uom_kind") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='uom_kind' size='5' value=''></td>"
				End If
			End If

			response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"

			Response.Write "<td id='mediumtd' style='text-align:center'><a href='unitsofmeasure.asp?sort=" & sortkey & "&direction=" & sortdir & "&limit=" & limitnum & "' title='Cancel changes to this record'>Cancel</a></td>"

			Response.Write "</tr>"
		End If
	End If

	Do While Not rs.EOF
		Response.Write "<tr>"
		If CLng(rs("uom_id")) = CLng(recordid) Then
			'Draw the data entry line
			response.write "<td id='mediumtd'>" & rs("uom_id") & "</td>"

			If session("err") = "uom_code" Then
				If session("uom_code") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='uom_code' size='10' value='" & session("uom_code") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='uom_code' size='10' value='" & rs("uom_code") & "'></td>"
				End If
			Else
				If session("uom_code") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='uom_code' size='10' value='" & session("uom_code") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='uom_code' size='10' value='" & rs("uom_code") & "'></td>"
				End If
			End If

			If session("err") = "uom_name" Then
				If session("uom_name") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='uom_name' size='20' value='" & session("uom_name") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='uom_name' size='20' value='" & rs("uom_name") & "'></td>"
				End If
			Else
				If session("uom_name") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='uom_name' size='20' value='" & session("uom_name") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='uom_name' size='20' value='" & rs("uom_name") & "'></td>"
				End If
			End If

			If session("err") = "uom_desc" Then
				If session("uom_desc") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='uom_desc' cols='30' rows='2'>" & session("uom_desc") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><textarea name='uom_desc' cols='30' rows='2'>" & rs("uom_desc") & "</textarea></td>"
				End If
			Else
				If session("uom_desc") <> "" Then
					response.write "<td id='mediumtd'><textarea name='uom_desc' cols='30' rows='2'>" & session("uom_desc") & "</textarea></td>"
				Else
					response.write "<td id='mediumtd'><textarea name='uom_desc' cols='30' rows='2'>" & rs("uom_desc") & "</textarea></td>"
				End If
			End If

			'Dropdown for unit of measure type.
			If session("err") = "uom_type" Then
				response.write "<td id='mediumtd' style='background-color:red;text-align:center'><select name='uom_type'>"
			Else
				response.write "<td id='mediumtd' style='text-align:center'><select name='uom_type'>"
			End If
			sqlString = "SELECT uom_type_id,CONCAT(uom_type_id,' - ',uom_type_desc) AS uom_type_name FROM unit_of_measure_types ORDER BY uom_type_id"
			Set rs2 = cn.Execute(sqlString)
			If Not rs2.BOF Then
				rs2.MoveFirst
				Response.Write "<option value=''> "
				Do While Not rs2.EOF
					If Request("uom_type") <> "" Then
						If CInt(Request("uom_type")) = rs2("uom_type_id") Then
							response.write "<option value='" & rs2("uom_type_id") & "' selected>" & rs2("uom_type_name")
						Else
							response.write "<option value='" & rs2("uom_type_id") & "'>" & rs2("uom_type_name")
						End If
					Else
						If rs("uom_type") = rs2("uom_type_id") Then
							response.write "<option value='" & rs2("uom_type_id") & "' selected>" & rs2("uom_type_name")
						Else
							response.write "<option value='" & rs2("uom_type_id") & "'>" & rs2("uom_type_name")
						End If
					End If
					rs2.MoveNext
				Loop
			End If
			rs2.Close
			response.write "</select></td>"

			If session("err") = "uom_kind" Then
				If session("uom_kind") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='uom_kind' size='5' value='" & session("uom_kind") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='uom_kind' size='5' value='" & rs("uom_kind") & "'></td>"
				End If
			Else
				If session("uom_kind") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='uom_kind' size='5' value='" & session("uom_kind") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='uom_kind' size='5' value='" & rs("uom_kind") & "'></td>"
				End If
			End If

			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"
			End If

			If access = "delete" Then
				recid = rs("uom_id")
				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
			End If
		Else
			'Draw the history records
			response.write "<tr>"
			response.write "<td id='mediumtd'>" & rs("uom_id") & "</td>"
			If Trim(rs("uom_code")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("uom_code") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("uom_name")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("uom_name") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("uom_desc")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("uom_desc") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("uom_type")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("uom_type") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("uom_kind")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("uom_kind") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd'><a href='unitsofmeasure.asp?record_id=" & rs("uom_id") & "&sort=" & sortkey & "&direction=" & sortdir & "&limit=" & limitnum & "' title='Edit this record'>Edit</a></td>"
			End If
			If access = "delete" Then
				recid = rs("uom_id")
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
