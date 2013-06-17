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
</script>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Administer Pipe Materials</title>
<link rel=STYLESHEET href='equipmentstyle.css' type='text/css'>
<style>
  input {font-family:verdana;
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
' Keith Brooks - Monday, May 3, 2010
'   Creation.
'*************

'on error resume next

dim cn
dim rs
dim recordid
Dim currentuser
Dim access
Dim recid
Dim sqlString
Dim sortkey
Dim sortdir

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("equipment","pipematerials",currentuser)
If access <> "none" Then

	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")

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
		sortkey = "pipe_mat_id"
	End If
	If request("direction") <> "" Then
		sortdir = request("direction")
	Else
		sortdir = "ASC"
	End If

	response.write "<table ID='headertable' width='100%'>"
	response.write "<tr>"
	response.write "<td ID='headertd' style='width:20%;text-align:left;vertical-align:top'><a href='adminmenu.asp' title='Open the administration main menu'>Menu</a></td>"
	response.write "<td ID='headertd' style='width:60%;text-align:center;vertical-align:center'><h1/>Edit Pipe Materials</td>"
	response.write "<td ID='headertd' style='width:20%;text-align:right;vertical-align:top'><a href='#' onClick='openhelp();return false;' title='Open the Admin Guide'>Help</a></td>"
	response.write "</tr>"

	response.write "<tr>"
	response.write "<td id='headertd'>&nbsp;</td>"
	If access = "write" Or access = "delete" Then
		response.write "<td ID='headertd' style='text-align:center'><a href='pipematerials.asp?record_id=-1&sort=" & sortkey & "&direction=" & sortdir & "' title='Add a new loop function record'>Add new record</a></td>"
	Else
		Response.Write "<td id='headertd'>&nbsp;</td>"
	End If
	response.write "<td id='headertd'>&nbsp;</td>"
	response.write "</tr>"
	response.write "</table>"

	Response.Write "<br />"

	'Draw the header
	Response.Write "<div style='text-align:center'>"
	response.Write "<table width='100%'>"
	response.Write "<tr>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=pipe_mat_id&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Pipe Mat.<br />ID&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=pipe_mat_id&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=pipe_mat_name&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Pipe Material Name&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=pipe_mat_name&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=linear_exp_coef&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Linear Exp.<br />Coefficient&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=linear_exp_coef&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=linear_exp_coef_add&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Linear Exp.<br />Coeff. Add.&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=linear_exp_coef_add&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=linear_exp_coef_uid&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Linear Exp.<br />Coeff. Units&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=linear_exp_coef_uid&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=border_temp&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Border<br />Temp.&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=border_temp&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=temp_min&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Temp.<br />Min.&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=temp_min&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=temp_max&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Temp.<br />Max.&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=temp_max&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th id='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=temp_uid&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Temp.<br />Units&nbsp;</th>"
	Response.Write "<td id='headertd'><a href='pipematerials.asp?sort=temp_uid&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	If access = "write" Or access = "delete" Then
 	  response.Write "<th id='mediumth'>&nbsp;</th>"
 	End If
	If access = "delete" Or recordid < 0 Then
 	  response.Write "<th id='mediumth'>&nbsp;</th>"
 	End If
	response.Write "</tr>"

	response.Write "<form action='adminaction.asp' id='form1' method='post' name='form1'>"

	response.write "<input type='hidden' name='RECORD' value='" & recordid & "'>"
	response.write "<input type='hidden' name='SORT' value='" & sortkey & "'>"
	response.write "<input type='hidden' name='DIRECTION' value='" & sortdir & "'>"

	sqlString = "SELECT * FROM pipe_materials " & _
				"ORDER BY " & sortkey & " " & sortdir
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

			If session("err") = "pipe_mat_name" Then
				If session("pipe_mat_name") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='pipe_mat_name' size='45' value='" & session("pipe_mat_name") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='pipe_mat_name' size='45' value=''></td>"
				End If
			Else
				If session("pipe_mat_name") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='pipe_mat_name' size='45' value='" & session("pipe_mat_name") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='pipe_mat_name' size='45' value=''></td>"
				End If
			End If

			If session("err") = "linear_exp_coef" Then
				If session("linear_exp_coef") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='linear_exp_coef' size='5' value='" & session("linear_exp_coef") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='linear_exp_coef' size='5' value=''></td>"
				End If
			Else
				If session("linear_exp_coef") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='linear_exp_coef' size='5' value='" & session("linear_exp_coef") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='linear_exp_coef' size='5' value=''></td>"
				End If
			End If

			If session("err") = "linear_exp_coef_add" Then
				If session("linear_exp_coef_add") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='linear_exp_coef_add' size='10' value='" & session("linear_exp_coef_add") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='linear_exp_coef_add' size='10' value=''></td>"
				End If
			Else
				If session("linear_exp_coef_add") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='linear_exp_coef_add' size='10' value='" & session("linear_exp_coef_add") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='linear_exp_coef_add' size='10' value=''></td>"
				End If
			End If

			If session("err") = "linear_exp_coef_uid" Then
				If session("linear_exp_coef_uid") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='linear_exp_coef_uid' size='5' value='" & session("linear_exp_coef_uid") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='linear_exp_coef_uid' size='5' value=''></td>"
				End If
			Else
				If session("linear_exp_coef_uid") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='linear_exp_coef_uid' size='5' value='" & session("linear_exp_coef_uid") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='linear_exp_coef_uid' size='5' value=''></td>"
				End If
			End If

			If session("err") = "border_temp" Then
				If session("border_temp") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='border_temp' size='5' value='" & session("border_temp") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='border_temp' size='5' value=''></td>"
				End If
			Else
				If session("border_temp") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='border_temp' size='5' value='" & session("border_temp") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='border_temp' size='5' value=''></td>"
				End If
			End If

			If session("err") = "temp_min" Then
				If session("temp_min") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='temp_min' size='5' value='" & session("temp_min") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='temp_min' size='5' value=''></td>"
				End If
			Else
				If session("temp_min") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='temp_min' size='5' value='" & session("temp_min") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='temp_min' size='5' value=''></td>"
				End If
			End If

			If session("err") = "temp_max" Then
				If session("temp_max") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='temp_max' size='5' value='" & session("temp_max") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='temp_max' size='5' value=''></td>"
				End If
			Else
				If session("temp_max") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='temp_max' size='5' value='" & session("temp_max") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='temp_max' size='5' value=''></td>"
				End If
			End If

			If session("err") = "temp_uid" Then
				If session("temp_uid") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='temp_uid' size='5' value='" & session("temp_uid") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='temp_uid' size='5' value=''></td>"
				End If
			Else
				If session("temp_uid") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='temp_uid' size='5' value='" & session("temp_uid") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='temp_uid' size='5' value=''></td>"
				End If
			End If

			response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"

			Response.Write "<td id='mediumtd' style='text-align:center'><a href='pipematerials.asp?sort=" & sortkey & "&direction=" & sortdir & "' title='Cancel changes to this record'>Cancel</a></td>"

			Response.Write "</tr>"
		End If
	End If

	Do While Not rs.EOF
		Response.Write "<tr>"
		If CLng(rs("pipe_mat_id")) = CLng(recordid) Then
			'Draw the data entry line
			response.write "<td id='mediumtd'>" & rs("pipe_mat_id") & "</td>"

			If session("err") = "pipe_mat_name" Then
				If session("pipe_mat_name") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='pipe_mat_name' size='45' value='" & session("pipe_mat_name") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='pipe_mat_name' size='45' value='" & rs("pipe_mat_name") & "'></td>"
				End If
			Else
				If session("pipe_mat_name") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='pipe_mat_name' size='45' value='" & session("pipe_mat_name") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='pipe_mat_name' size='45' value='" & rs("pipe_mat_name") & "'></td>"
				End If
			End If

			If session("err") = "linear_exp_coef" Then
				If session("linear_exp_coef") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='linear_exp_coef' size='5' value='" & session("linear_exp_coef") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='linear_exp_coef' size='5' value='" & rs("linear_exp_coef") & "'></td>"
				End If
			Else
				If session("linear_exp_coef") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='linear_exp_coef' size='5' value='" & session("linear_exp_coef") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='linear_exp_coef' size='5' value='" & rs("linear_exp_coef") & "'></td>"
				End If
			End If

			If session("err") = "linear_exp_coef_add" Then
				If session("linear_exp_coef_add") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='linear_exp_coef_add' size='10' value='" & session("linear_exp_coef_add") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='linear_exp_coef_add' size='10' value='" & rs("linear_exp_coef_add") & "'></td>"
				End If
			Else
				If session("linear_exp_coef_add") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='linear_exp_coef_add' size='10' value='" & session("linear_exp_coef_add") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='linear_exp_coef_add' size='10' value='" & rs("linear_exp_coef_add") & "'></td>"
				End If
			End If

			If session("err") = "linear_exp_coef_uid" Then
				If session("linear_exp_coef_uid") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='linear_exp_coef_uid' size='5' value='" & session("linear_exp_coef_uid") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='linear_exp_coef_uid' size='5' value='" & rs("linear_exp_coef_uid") & "'></td>"
				End If
			Else
				If session("linear_exp_coef_uid") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='linear_exp_coef_uid' size='5' value='" & session("linear_exp_coef_uid") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='linear_exp_coef_uid' size='5' value='" & rs("linear_exp_coef_uid") & "'></td>"
				End If
			End If

			If session("err") = "border_temp" Then
				If session("border_temp") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='border_temp' size='5' value='" & session("border_temp") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='border_temp' size='5' value='" & rs("border_temp") & "'></td>"
				End If
			Else
				If session("border_temp") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='border_temp' size='5' value='" & session("border_temp") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='border_temp' size='5' value='" & rs("border_temp") & "'></td>"
				End If
			End If

			If session("err") = "temp_min" Then
				If session("temp_min") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='temp_min' size='5' value='" & session("temp_min") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='temp_min' size='5' value='" & rs("temp_min") & "'></td>"
				End If
			Else
				If session("temp_min") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='temp_min' size='5' value='" & session("temp_min") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='temp_min' size='5' value='" & rs("temp_min") & "'></td>"
				End If
			End If

			If session("err") = "temp_max" Then
				If session("temp_max") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='temp_max' size='5' value='" & session("temp_max") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='temp_max' size='5' value='" & rs("temp_max") & "'></td>"
				End If
			Else
				If session("temp_max") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='temp_max' size='5' value='" & session("temp_max") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='temp_max' size='5' value='" & rs("temp_max") & "'></td>"
				End If
			End If

			If session("err") = "temp_uid" Then
				If session("temp_uid") <> "" Then
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='temp_uid' size='5' value='" & session("temp_uid") & "'></td>"
				Else
					response.write "<td id='mediumtd' style='background-color:red'><input type='text' name='temp_uid' size='5' value='" & rs("temp_uid") & "'></td>"
				End If
			Else
				If session("temp_uid") <> "" Then
					response.write "<td id='mediumtd'><input type='text' name='temp_uid' size='5' value='" & session("temp_uid") & "'></td>"
				Else
					response.write "<td id='mediumtd'><input type='text' name='temp_uid' size='5' value='" & rs("temp_uid") & "'></td>"
				End If
			End If

			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"
			End If

			If access = "delete" Then
				recid = rs("pipe_mat_id")
				response.write "<td id='mediumtd'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
			End If
		Else
			'Draw the history records
			response.write "<tr>"
			response.write "<td id='mediumtd'>" & rs("pipe_mat_id") & "</td>"
			If rs("pipe_mat_name") <> "" Then
				response.write "<td id='mediumtd'>" & rs("pipe_mat_name") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("linear_exp_coef")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("linear_exp_coef") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("linear_exp_coef_add")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("linear_exp_coef_add") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("linear_exp_coef_uid")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("linear_exp_coef_uid") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("border_temp")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("border_temp") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("temp_min")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("temp_min") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("temp_max")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("temp_max") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If Trim(rs("temp_uid")) <> "" Then
				response.write "<td id='mediumtd'>" & rs("temp_uid") & "</td>"
			Else
				Response.Write "<td id='mediumtd'>&nbsp;</td>"
			End If
			If access = "write" Or access = "delete" Then
				response.write "<td id='mediumtd'><a href='pipematerials.asp?record_id=" & rs("pipe_mat_id") & "&sort=" & sortkey & "&direction=" & sortdir & "' title='Edit this record'>Edit</a></td>"
			End If
			If access = "delete" Then
				recid = rs("pipe_mat_id")
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
