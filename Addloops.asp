<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="javascript">
function doSubmit() {
 document.getElementById('PleaseWait').style.display = 'block';
 document.form1.submit();
}
function addItem() {
 if (document.form1.plant_area_id.value == '') {
  alert('You must select a plant area');
 } else {
  if (document.form1.loop_proc_id.value == '') {
   alert('You must select a loop process');
  } else {
   if (document.form1.loop_num.value.length < 5) {
    alert('You must enter a 5-digit loop number');
   } else {
    document.getElementById('PleaseWait').style.display = 'block';
    document.form1.action='adminaction.asp?RECORD=-1';
    document.form1.submit();
   }
  }
 }
}
function copyItem() {
 if (document.form1.plantarea.value == '') {
  alert('You must select a plant area');
 } else {
  if (document.form1.instrumenttype.value == '') {
   alert('You must select an instrument type');
  } else {
   if (document.form1.instrnumber.value == '') {
    alert('You must enter an instrument number');
   } else {
    document.getElementById('PleaseWait').style.display = 'block';
    document.form1.action='updatespec.asp?copy=true';
    document.form1.submit();
   }
  }
 }
}
function openhelp() {
 window.open("Instrument Spec Database Administrators Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Add Loop</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="EquipmentFunctions.asp"-->
<link rel=STYLESHEET href='equipmentstyle.css' type='text/css'>
<style>
  div {font-family:verdana;}
  input {font-family:verdana;}
  select {font-family:verdana;
		font-size:10pt;}
  textarea {font-family:verdana;}
</style>
</head>
<body>
<table style="width:100%;border:none">
	<tr>
		<td style="text-align:left;vertical-align:top;width:20%" title="Open the administration main menu"><a href="adminmenu.asp">Menu</a></td>
		<td style="text-align:center;width:60%"><h1 />Add Loop</td>
		<td style="text-align:right;vertical-align:top;width:20%"><a href="" onclick="openhelp();return false;" title="Open the Admin Guide">Help</a></td>
	</tr>
</table>
<form id="form1" name="form1" action="addloops.asp" method="post">
<%
Dim sqlString
Dim cn
Dim rs
Dim criteria
Dim tagname
Dim equipnum
Dim currentuser
Dim access
Dim access2

'Define the ado connection and recordset objects.
set cn = CreateObject("adodb.connection")
cn.Open = DBString
set rs = CreateObject("adodb.recordset")

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
'access = UserAllowed(currentuser, "masterbatchentry")
access = UserAccess("equipment", "addloops", currentuser)
If access <> "none" Then

	'Draw "Please Wait..." message that will be displayed when this page is
	'reloading, saving data, or moving to another page.
	%>
		<div class="helptext" id="PleaseWait" style="display: none; text-align:center; color:White; vertical-align:top;border-style:none;position:absolute;top:0px;left:0px">
			<table id="MyTable" bgcolor="blue">
				<tr><td style="width: 95px"><b><font color="white">Please Wait...</font></b></td></tr>
			</table>
		</div>
	<%
	'Load the plant area dropdown list.
	Response.Write "<div style='text-align:center'>"
	Response.Write "<table style='width:70%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<th style='width:30%;text-align:right'>Plant Area:&nbsp;&nbsp;</th>"
	Response.Write "<td style='text-align:left'><select id='plant_area_id' name='plant_area_id' onchange='doSubmit();'>"
	sqlString = "SELECT plant_area_id,plant_area_name FROM plant_areas ORDER BY plant_area_name"
	Set rs = cn.Execute(sqlString)
	If Not rs.BOF Then
		rs.MoveFirst
		Response.Write "<option value=''> "
		Do While Not rs.EOF
			If Request("plant_area_id") <> "" Then
				If CLng(rs(0)) = CLng(Request("plant_area_id")) Then
					Response.Write "<option value='" & rs(0) & "' selected>" & rs(1)
				Else
					Response.Write "<option value='" & rs(0) & "'>" & rs(1)
				End If
			Else
				Response.Write "<option value='" & rs(0) & "'>" & rs(1)
			End If
			rs.MoveNext
		Loop
	End If
	rs.Close
	Response.Write "</select></td>"
	Response.Write "</tr>"

	If Request("plant_area_id") <> "" Then

		'Load the equipment dropdown list.
		Response.Write "<tr>"
		Response.Write "<th style='text-align:right'>Equipment:&nbsp;&nbsp;</th>"
		Response.Write "<td style='text-align:left'><select id='loop_equip_id' name='loop_equip_id' onchange='doSubmit();'>"
		If Request("plant_area_id") <> "" Then
			sqlString = "SELECT equip_id,equip_name FROM equipment WHERE plant_area_id=" & Request("plant_area_id") & " ORDER BY equip_name"
		Else
			sqlString = "SELECT equip_id,equip_name FROM equipment ORDER BY equip_name"
		End If
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If Request("loop_equip_id") <> "" Then
					If CLng(rs(0)) = CLng(Request("loop_equip_id")) Then
						Response.Write "<option value='" & rs(0) & "' selected>" & rs(1)
					Else
						Response.Write "<option value='" & rs(0) & "'>" & rs(1)
					End If
				Else
					Response.Write "<option value='" & rs(0) & "'>" & rs(1)
				End If
				rs.MoveNext
			Loop
		End If
		rs.Close
		Response.Write "</select></td>"
		Response.Write "</tr>"

		'Load the Loop Process dropdown list.
		Response.Write "<tr>"
		Response.Write "<th style='text-align:right'>Loop Process:&nbsp;&nbsp;</th>"
		Response.Write "<td style='text-align:left'><select name='loop_proc_id' id='loop_proc_id' onchange='doSubmit();'>"
		sqlString = "SELECT loop_proc_id,CONCAT(loop_proc_name,' - ',loop_proc_desc) AS loop_proc FROM loop_processes ORDER BY loop_proc"
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If Request("loop_proc_id") <> "" Then
					If CLng(rs(0)) = CLng(Request("loop_proc_id")) Then
						Response.Write "<option value='" & rs(0) & "' selected>" & rs(1)
					Else
						Response.Write "<option value='" & rs(0) & "'>" & rs(1)
					End If
				Else
					Response.Write "<option value='" & rs(0) & "'>" & rs(1)
				End If
				rs.MoveNext
			Loop
		End If
		rs.Close
		Response.Write "</select></td>"
		Response.Write "</tr>"

		'Load the Loop Type dropdown list.
		Response.Write "<tr>"
		Response.Write "<th style='text-align:right'>Loop Type:&nbsp;&nbsp;</th>"
		Response.Write "<td style='text-align:left'><select name='loop_type_id' id='loop_type_id'>"
		sqlString = "SELECT loop_type_id,loop_type_name FROM loop_types ORDER BY loop_type_name"
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If Request("loop_type_id") <> "" Then
					If CLng(rs(0)) = CLng(Request("loop_type_id")) Then
						Response.Write "<option value='" & rs(0) & "' selected>" & rs(1)
					Else
						Response.Write "<option value='" & rs(0) & "'>" & rs(1)
					End If
				Else
					Response.Write "<option value='" & rs(0) & "'>" & rs(1)
				End If
				rs.MoveNext
			Loop
		End If
		rs.Close
		Response.Write "</select></td>"
		Response.Write "</tr>"
		
		'Load the Loop Function dropdown list.
		Response.Write "<tr>"
		Response.Write "<th style='text-align:right'>Loop Function:&nbsp;&nbsp;</th>"
		Response.Write "<td style='text-align:left'><select name='loop_func_id' id='loop_func_id'>"
		sqlString = "SELECT loop_func_id,loop_func_name FROM loop_functions ORDER BY loop_func_name"
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If Request("loop_func_id") <> "" Then
					If CLng(rs(0)) = CLng(Request("loop_func_id")) Then
						Response.Write "<option value='" & rs(0) & "' selected>" & rs(1)
					Else
						Response.Write "<option value='" & rs(0) & "'>" & rs(1)
					End If
				Else
					Response.Write "<option value='" & rs(0) & "'>" & rs(1)
				End If
				rs.MoveNext
			Loop
		End If
		rs.Close
		Response.Write "</select></td>"
		Response.Write "</tr>"
		
		'Load the Drawing dropdown list.
		Response.Write "<tr>"
		Response.Write "<th style='text-align:right'>Drawing:&nbsp;&nbsp;</th>"
		Response.Write "<td style='text-align:left'><select name='dwg_id' id='dwg_id'>"
		sqlString = "SELECT dwg_id,dwg_name FROM drawings WHERE dwg_name is not null and dwg_type_id=1 ORDER BY dwg_name"
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If Request("dwg_id") <> "" Then
					If CLng(rs(0)) = CLng(Request("dwg_id")) Then
						Response.Write "<option value='" & rs(0) & "' selected>" & rs(1)
					Else
						Response.Write "<option value='" & rs(0) & "'>" & rs(1)
					End If
				Else
					Response.Write "<option value='" & rs(0) & "'>" & rs(1)
				End If
				rs.MoveNext
			Loop
		End If
		rs.Close
		Response.Write "</select></td>"
		Response.Write "</tr>"
		
		'Draw the loop number textbox.
		Response.Write "<tr>"
		Response.Write "<th style='text-align:right'>Loop Number:&nbsp;&nbsp;</th>"
		If Request("loop_num") <> "" Then
			Response.Write "<td style='text-align:left'><input type='textbox' id='loop_num' name='loop_num' size='6' value='" & Request("loop_num") & "' onchange='doSubmit();' /></td>"
		Else
			equipnum = ""
			If Request("loop_equip_id") <> "" Then
				sqlString = "SELECT SUBSTRING(equip_name,-3) FROM equipment WHERE equip_id=" & Request("loop_equip_id")
				Set rs = cn.Execute(sqlString)
				If Not rs.BOF Then
					rs.MoveFirst
					If IsNumeric(rs(0)) Then
						equipnum = rs(0)
					End If
				End If
			End If
			Response.Write "<td style='text-align:left'><input type='textbox' id='loop_num' name='loop_num' size='6' value='" & equipnum & "' onchange='doSubmit();' /></td>"
		End If
		Response.Write "</tr>"
		
		'Draw the suffix textbox.
		Response.Write "<tr>"
		Response.Write "<th style='text-align:right'>Loop Suffix:&nbsp;&nbsp;</th>"
		Response.Write "<td style='text-align:left'><input type='textbox' id='loop_suff' name='loop_suff' size='2' value='" & Request("loop_suff") & "' onchange='doSubmit();' /></td>"
		Response.Write "</tr>"
		
		'Draw the loop name text box.  Automatically fill it in as data is entered above.
		Response.Write "<tr>"
		Response.Write "<th style='text-align:right'>Loop Name:&nbsp;&nbsp;</th>"
		tagname = ""
		If Request("loop_proc_id") <> "" Then
			sqlString = "SELECT loop_proc_name FROM loop_processes WHERE loop_proc_id=" & Request("loop_proc_id")
			Set rs = cn.Execute(sqlString)
			If Not rs.BOF Then
				rs.MoveFirst
				If Not IsNull(rs(0)) Then
					tagname = PadRight(Trim(rs(0)),4)
				Else
					tagname = "    "
				End If
			Else
				tagname = "    "
			End If
			rs.Close
		Else
			tagname = "    "
		End If
		If Request("plant_area_id") <> "" Then
			sqlString = "SELECT plant_area_num FROM plant_areas WHERE plant_area_id=" & Request("plant_area_id")
			Set rs = cn.Execute(sqlString)
			If Not rs.BOF Then
				rs.MoveFirst
				If Not IsNull(rs(0)) Then
					tagname = tagname & "-" & Trim(rs(0)) & "-"
				Else
					tagname = tagname & "-    -"
				End If
			Else
				tagname = tagname & "-    -"
			End If
			rs.Close
		Else
			tagname = tagname & "-    -"
		End If
		If Request("loop_num") <> "" Then
			tagname = tagname & Trim(Request("loop_num"))
		Else
			If Request("loop_equip_id") <> "" Then
				tagname = tagname & equipnum
			End If
		End If
		If Request("loop_suff") <> "" Then
			tagname = tagname & Trim(Request("loop_suff"))
		End If
		Response.Write "<td style='text-align:left'><input type='textbox' id='loop_name' name='loop_name' size='20' style='font-family:courier new' value='" & Replace(tagname," ","&nbsp;") & "' /></td>"
		Response.Write "</tr>"
		
		'Draw the description textarea.
		Response.Write "<tr>"
		Response.Write "<th style='text-align:right'>Loop Description:&nbsp;&nbsp;</th>"
		Response.Write "<td style='text-align:left'><textarea id='loop_desc' name='loop_desc' cols='65' rows='4'>" & Request("loop_desc") & "</textarea></td>"
		Response.Write "</tr>"

		'Draw the note textarea.
		Response.Write "<tr>"
		Response.Write "<th style='text-align:right'>Loop Note:&nbsp;&nbsp;</th>"
		Response.Write "<td style='text-align:left'><textarea id='loop_note' name='loop_note' cols='65' rows='4'>" & Request("loop_note") & "</textarea></td>"
		Response.Write "</tr>"

	End If
	Response.Write "</table>"
	Response.Write "</div>"

	If access = "write" Or access = "delete" Then
		If Request("plant_area_id") <> "" Then
			Response.Write "<br />"
			Response.Write "<div style='text-align:center'>"
			Response.Write "<table style='width:40%'>"
			Response.Write "<tr>"
			Response.Write "<td style='width:100%;text-align:center'><input type='button' id='submit1' name='submit1' value='Add' style='font-size:10pt' title='Add this instrument and select a blank spec form to edit' onclick='addItem();'></td>"
'			Response.Write "<td style='width:50%;text-align:center'><input type='button' id='copy1' name='copy1' value='Copy' style='font-size:10pt' title='Add this intrument and select an existing instrument spec to copy' onclick='copyItem();'></td>"
			Response.Write "</tr>"
			Response.Write "</table>"
			Response.Write "</div>"
		End If
	End If

Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If

Set rs = Nothing
cn.Close
Set cn = Nothing
%>
</form>
</body>
</html>