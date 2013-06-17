<%@ language="vbscript" %>
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
 if (document.form1.plantarea.value == '') {
  alert('You must select a plant area');
 } else {
  if (document.form1.instrumenttype.value == '') {
   alert('You must select an instrument type');
  } else {
   if (document.form1.instrnumber.value.length < 5) {
    alert('You must enter a 5-digit instrument number');
   } else {
    document.getElementById('PleaseWait').style.display = 'block';
    document.form1.action='updatespec.asp';
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
 window.open("Instrument Spec Database Users Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Add Instrument</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="EquipmentFunctions.asp"-->
<link rel="stylesheet" href="equipmentstyle.css" type="text/css">
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
		<td class="left top" style="width:20%"><a class="noprint" href="default.asp">Home</a></td>
		<td class="center" style="width:60%"><h1 />Add Instrument</td>
		<td class="right top" style="width:20%"><a class="noprint" href="#" onclick="openhelp();return false;" title="Open the User Guide">Help</a></td>
	</tr>
</table>
<form id="form1" name="form1" action="AddInstrument.asp" method="post">
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
access = UserAccess("equipment", "addinstrument", currentuser)
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
	Response.Write "<div class='center'>"
	Response.Write "<table style='width:70%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<th class='right' style='width:30%'>Plant Area:&nbsp;&nbsp;</th>"
	Response.Write "<td class='left'><select id='plantarea' name='plantarea' onchange='doSubmit();'>"
	sqlString = "SELECT plant_area_id,plant_area_name FROM plant_areas ORDER BY plant_area_name"
	Set rs = cn.Execute(sqlString)
	If Not rs.BOF Then
		rs.MoveFirst
		Response.Write "<option value=''> "
		Do While Not rs.EOF
			If Request("plantarea") <> "" Then
				If CLng(rs(0)) = CLng(Request("plantarea")) Then
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

	If Request("plantarea") <> "" Then

		'Load the equipment dropdown list.
		Response.Write "<tr>"
		Response.Write "<th class='right'>Equipment:&nbsp;&nbsp;</th>"
		Response.Write "<td class='left'><select id='equipment' name='equipment' onchange='doSubmit();'>"
		If Request("plantarea") <> "" Then
			sqlString = "SELECT equip_id,equip_name FROM equipment WHERE plant_area_id=" & Request("plantarea") & " ORDER BY equip_name"
		Else
			sqlString = "SELECT equip_id,equip_name FROM equipment ORDER BY equip_name"
		End If
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If Request("equipment") <> "" Then
					If CLng(rs(0)) = CLng(Request("equipment")) Then
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

		'Load the Loop dropdown list.
		Response.Write "<tr>"
		Response.Write "<th class='right'>Loop:&nbsp;&nbsp;</th>"
		Response.Write "<td class='left'><select name='loop' id='loop'>"
		If Request("plantarea") <> "" Then
			sqlString = "SELECT loop_id,loop_name,loop_desc FROM loops WHERE plant_area_id=" & Request("plantarea") & " ORDER BY loop_name"
		Else
			sqlString = "SELECT loop_id,loop_name,loop_desc FROM loops ORDER BY loop_name"
		End If
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If Request("loop") <> "" Then
					If CLng(rs(0)) = CLng(Request("loop")) Then
						Response.Write "<option value='" & rs(0) & "' selected>" & rs(1) & " - " & rs(2)
					Else
						Response.Write "<option value='" & rs(0) & "'>" & rs(1) & " - " & rs(2)
					End If
				Else
					Response.Write "<option value='" & rs(0) & "'>" & rs(1) & " - " & rs(2)
				End If
				rs.MoveNext
			Loop
		End If
		rs.Close
		Response.Write "</select></td>"
		Response.Write "</tr>"

		'Load the Instrument Type dropdown list.
		Response.Write "<tr>"
		Response.Write "<th class='right'>Instrument Type:&nbsp;&nbsp;</th>"
		Response.Write "<td class='left'><select name='instrumenttype' id='instrumenttype' onchange='doSubmit();'>"
		If Request("processfunction") <> "" Then
			sqlString = "SELECT instr_func_type_id,instr_func_type_name,instr_func_type_desc FROM instrument_function_types WHERE proc_func_id=" & Request("processfunction") & " ORDER BY instr_func_type_name"
		Else
			sqlString = "SELECT instr_func_type_id,instr_func_type_name,instr_func_type_desc FROM instrument_function_types ORDER BY instr_func_type_name"
		End If
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If Request("instrumenttype") <> "" Then
					If CLng(rs(0)) = CLng(Request("instrumenttype")) Then
						Response.Write "<option value='" & rs(0) & "' selected>" & rs(1) & " - " & rs(2)
					Else
						Response.Write "<option value='" & rs(0) & "'>" & rs(1) & " - " & rs(2)
					End If
				Else
					Response.Write "<option value='" & rs(0) & "'>" & rs(1) & " - " & rs(2)
				End If
				rs.MoveNext
			Loop
		End If
		rs.Close
		Response.Write "</select></td>"
		Response.Write "</tr>"
		
		'Draw the instrument number textbox.
		Response.Write "<tr>"
		Response.Write "<th class='right'>Instrument Number:&nbsp;&nbsp;</th>"
		If Request("instrnumber") <> "" Then
			Response.Write "<td class='left'><input type='textbox' id='instrnumber' name='instrnumber' size='6' value='" & Request("instrnumber") & "' onchange='doSubmit();' /></td>"
		Else
			equipnum = ""
			If Request("equipment") <> "" Then
				sqlString = "SELECT SUBSTRING(equip_name,-3) FROM equipment WHERE equip_id=" & Request("equipment")
				Set rs = cn.Execute(sqlString)
				If Not rs.BOF Then
					rs.MoveFirst
					If IsNumeric(rs(0)) Then
						equipnum = rs(0)
					End If
				End If
			End If
			Response.Write "<td class='left'><input type='textbox' id='instrnumber' name='instrnumber' size='6' value='" & equipnum & "' onchange='doSubmit();' /></td>"
		End If
		Response.Write "</tr>"
		
		'Draw the suffix textbox.
		Response.Write "<tr>"
		Response.Write "<th class='right'>Tag Suffix:&nbsp;&nbsp;</th>"
		Response.Write "<td class='left'><input type='textbox' id='suffix' name='suffix' size='2' value='" & Request("suffix") & "' onchange='doSubmit();' /></td>"
		Response.Write "</tr>"
		
		'Draw the tagname text box.  Automatically fill it in as data is entered above.
		Response.Write "<tr>"
		Response.Write "<th class='right'>Tag Number:&nbsp;&nbsp;</th>"
		tagname = ""
		If Request("instrumenttype") <> "" Then
			sqlString = "SELECT instr_func_type_name FROM instrument_function_types WHERE instr_func_type_id=" & Request("instrumenttype")
			Set rs = cn.Execute(sqlString)
			If Not rs.BOF Then
				rs.MoveFirst
				If Not IsNull(rs(0)) Then
					tagname = PadRight(Trim(rs(0)),5)
				Else
					tagname = "     "
				End If
			Else
				tagname = "     "
			End If
			rs.Close
		Else
			tagname = "     "
		End If
		If Request("plantarea") <> "" Then
			sqlString = "SELECT plant_area_num FROM plant_areas WHERE plant_area_id=" & Request("plantarea")
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
		If Request("instrnumber") <> "" Then
			tagname = tagname & Trim(Request("instrnumber"))
		Else
			If Request("equipment") <> "" Then
				tagname = tagname & equipnum
			End If
		End If
		If Request("suffix") <> "" Then
			tagname = tagname & Trim(Request("suffix"))
		End If
		Response.Write "<td class='left'><input type='textbox' id='tagnumber' name='tagnumber' size='20' style='font-family:courier new' value='" & Replace(tagname," ","&nbsp;") & "' /></td>"
		Response.Write "</tr>"
		
		'Draw the remarks textarea.
		Response.Write "<tr>"
		Response.Write "<th class='right'>Remarks:&nbsp;&nbsp;</th>"
		Response.Write "<td class='left'><textarea id='remarks' name='remarks' cols='65' rows='4'>" & Request("remarks") & "</textarea></td>"
		Response.Write "</tr>"

	End If
	Response.Write "</table>"
	Response.Write "</div>"

	If access = "write" Or access = "delete" Then
		If Request("plantarea") <> "" Then
			Response.Write "<br />"
			Response.Write "<div class='center noprint'>"
			Response.Write "<table style='width:40%'>"
			Response.Write "<tr>"
			Response.Write "<td class='center' style='width:50%'><input type='button' id='submit1' name='submit1' value='Add' style='font-size:10pt' title='Add this instrument and select a blank spec form to edit' onclick='addItem();'></td>"
			Response.Write "<td class='center' style='width:50%'><input type='button' id='copy1' name='copy1' value='Copy' style='font-size:10pt' title='Add this intrument and select an existing instrument spec to copy' onclick='copyItem();'></td>"
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