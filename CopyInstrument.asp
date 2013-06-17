<%@ language="vbscript" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="javascript">
function doSubmit() {
 document.getElementById('PleaseWait').style.display = 'block';
 document.form1.flowflag.value='false';
 document.form1.submit();
}
function doFind() {
 document.getElementById('PleaseWait').style.display = 'block';
 document.form1.submit();
}
function copyspec(newid,copyid) {
 if (newid=='') {
  alert('You must select an instrument to copy.');
 } else {
  document.getElementById('PleaseWait').style.display = 'block';
  document.form1.action="updatespec.asp?newid="+newid+"&copyid="+copyid;
  document.form1.submit();
 }
}
function openhelp() {
 window.open("Instrument Spec Database Users Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Copy Instrument</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="EquipmentFunctions.asp"-->
<link rel="stylesheet" href="equipmentstyle.css" type="text/css">
<style>
  div {font-family:verdana;}
  input {font-family:verdana;}
  select {font-family:verdana;
		font-size:10pt;
		width:99%;}
  textarea {font-family:verdana;}
</style>
</head>
<body>
<table style="width:100%;border:none">
	<tr>
		<td class="left top" style="width:20%"><a class="noprint" href="default.asp">Home</a></td>
		<td class="center" style="width:60%"><h1 />Select Instrument Spec to Copy</td>
		<td class="right top" style="width:20%"><a class="noprint" href="#" onclick="openhelp();return false;" title="Open the User Guide">Help</a></td>
	</tr>
</table>
<form id="form1" name="form1" action="CopyInstrument.asp" method="post">
<%
Dim sqlString
Dim cn
Dim rs
Dim criteria
Dim tagname
Dim tagdesc
Dim currentuser
Dim access

'Define the ado connection and recordset objects.
set cn = CreateObject("adodb.connection")
cn.Open = DBString
set rs = CreateObject("adodb.recordset")

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("equipment","copyinstrument",currentuser)
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
	'Save the ID of the instrument being copied.
	Response.Write "<input type='hidden' id='instrID' name='instrID' value='" & Request("instrID") & "' />"

	'Draw the criteria selection lists.
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<th style='width:20%'>Plant Area</th>"
	Response.Write "<th style='width:25%'>Process Function</th>"
	Response.Write "<th style='width:30%'>Instrument Type</th>"
	Response.Write "<th style='width:25%'>Show Additional Filters</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"

	'Load the Process Areas dropdown list.
	Response.Write "<td><select name='plantarea' onchange='doSubmit();'>"
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
	'Load the Process Functions dropdown list.
	Response.Write "<td><select name='processfunction' onchange='doSubmit();'>"
	sqlString = "SELECT proc_func_id,proc_func_name FROM process_functions ORDER BY proc_func_name"
	Set rs = cn.Execute(sqlString)
	If Not rs.BOF Then
		rs.MoveFirst
		Response.Write "<option value=''> "
		Do While Not rs.EOF
			If Request("processfunction") <> "" Then
				If CLng(rs(0)) = CLng(Request("processfunction")) Then
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
	'Load the Instrument Type dropdown list.
	Response.Write "<td><select name='instrumenttype'>"
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
	If Request("showmore") <> "" Then
		Response.Write "<td><input type='checkbox' id='showmore' name='showmore' onclick='doSubmit();' checked /></td>"
	Else
		Response.Write "<td><input type='checkbox' id='showmore' name='showmore' onclick='doSubmit();' /></td>"
	End If
	Response.Write "</tr>"
	Response.Write "</table>"

	If Request("showmore") <> "" Then
		Response.Write "<table style='width:100%'>"
		Response.Write "<tr>"
		Response.Write "<th style='width:17%'>Equipment</th>"
		Response.Write "<th style='width:18%'>Line</th>"
		Response.Write "<th style='width:40%'>Loop</th>"
		Response.Write "<th style='width:25%'>Location</th>"
		Response.Write "</tr>"

		Response.Write "<tr>"
		'Load the Equipment dropdown list.
		Response.Write "<td><select name='equipment'>"
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
		'Load the Line dropdown list.
		Response.Write "<td><select name='line'>"
		If Request("plantarea") <> "" Then
			sqlString = "SELECT line_id,line_num FROM process_lines WHERE plant_area_id=" & Request("plantarea") & " ORDER BY line_num"
		Else
			sqlString = "SELECT line_id,line_num FROM process_lines ORDER BY line_num"
		End If
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If Request("line") <> "" Then
					If CLng(rs(0)) = CLng(Request("line")) Then
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
		'Load the Loop dropdown list.
		Response.Write "<td><select name='loop'>"
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
		'Load the Location dropdown list.
		Response.Write "<td><select name='location'>"
		If Request("plantarea") <> "" Then
			sqlString = "SELECT instr_location_id,instr_location_name FROM instrument_locations WHERE plant_area_id=" & Request("plantarea") & " ORDER BY instr_location_name"
		Else
			sqlString = "SELECT instr_location_id,instr_location_name FROM instrument_locations ORDER BY instr_location_name"
		End If
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If Request("location") <> "" Then
					If CLng(rs(0)) = CLng(Request("location")) Then
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
		Response.Write "</table>"
	End If
	Response.Write "<br />"
	Response.Write "<table style='width:100%'>"
	Response.Write "<tr>"
	Response.Write "<td style='width:33%'>&nbsp;</td>"
	Response.Write "<td class='center' style='width:34%'><input type='button' id='submit1' name='submit1' value='Find' style='font-size:10pt' onclick='doFind();'></td>"
	If Request("showall") <> "" Then
		Response.Write "<td class='right' style='width:33%'><input type='checkbox' id='showall' name='showall' onclick='doFind();' checked />Show Voided Instruments</td>"
	Else
		Response.Write "<td class='right' style='width:33%'><input type='checkbox' id='showall' name='showall' onclick='doFind();' />Show Voided Instruments</td>"
	End If
	Response.Write "</tr>"
	Response.Write "</table>"

	'If any of the criteria have been selected, display the list box with the results.
	criteria = "spec_id is not null"
	If Request("plantarea") <> "" Then
		If criteria = "" Then
			criteria = "plant_area_id=" & Request("plantarea")
		Else
			criteria = criteria & " AND plant_area_id=" & Request("plantarea")
		End If
	End If
	If Request("processfunction") <> "" Then
		If criteria = "" Then
			criteria = "proc_func_id=" & Request("processfunction")
		Else
			criteria = criteria & " AND proc_func_id=" & Request("processfunction")
		End If
	End If
	If Request("instrumenttype") <> "" Then
		If criteria = "" Then
			criteria = "instr_func_type_id=" & Request("instrumenttype")
		Else
			criteria = criteria & " AND instr_func_type_id=" & Request("instrumenttype")
		End If
	End If
	If Request("equipment") <> "" Then
		If criteria = "" Then
			criteria = "equip_id=" & Request("equipment")
		Else
			criteria = criteria & " AND equip_id=" & Request("equipment")
		End If
	End If
	If Request("line") <> "" Then
		If criteria = "" Then
			criteria = "line_id=" & Request("line")
		Else
			criteria = criteria & " AND line_id=" & Request("line")
		End If
	End If
	If Request("loop") <> "" Then
		If criteria = "" Then
			criteria = "loop_id=" & Request("loop")
		Else
			criteria = criteria & " AND loop_id=" & Request("loop")
		End If
	End If
	If Request("location") <> "" Then
		If criteria = "" Then
			criteria = "instr_location_id=" & Request("location")
		Else
			criteria = criteria & " AND instr_location_id=" & Request("location")
		End If
	End If
	If Request("showall") = "" Then
		If criteria <> "" Then
			criteria = criteria & " AND void_sequence=0"
		End If
	End If
	If Request("flowflag") = "true" And criteria <> "" Then
		sqlString = "SELECT instr_id,instr_name,instr_desc,spec_form_num FROM instruments LEFT JOIN spec_forms ON instruments.spec_id=spec_forms.spec_form_id WHERE " & criteria & " ORDER BY instr_name"
		Set rs = cn.Execute(sqlString)
		Response.Write "<div class='center'><table style='width:700px'>"
		Response.Write "<tr>"
		Response.Write "<th width='20%'>Instrument Name</th>"
		Response.Write "<th width='60%'>Instrument Description</th>"
		Response.Write "<th width='20%'>Spec Form</th>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td colspan='3'><select id='instr_id' size='17' style='font-family:courier new' onDblClick='copyspec(" & Request("instrID") & ",this.value);'>"
		If Not rs.BOF Then
			rs.MoveFirst
			Do While Not rs.EOF
				If Not IsNull(rs(1)) Then
					tagname = Replace(PadRight(rs(1),17)," ","&nbsp;")
				Else
					tagname = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				End If
				If Not IsNull(rs(2)) Then
					tagdesc = Replace(PadRight(rs(2),52)," ","&nbsp;")
				Else
					tagdesc = "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
							"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;" & _
							"&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;"
				End If
				Response.Write "<option value='" & rs(0) & "'>" & tagname & "&nbsp;&nbsp;&nbsp;&nbsp; " & tagdesc & "&nbsp;&nbsp; " & rs(3)
				rs.MoveNext
			Loop
		End If
		rs.Close
		Response.Write "</select></td>"
		Response.Write "</tr>"
		Response.Write "<tr class='noprint'>"
		Response.Write "<td>&nbsp;</td>"
		Response.Write "<td><button type='button' name='copybutton' id='copybutton' onclick='copyspec(" & Request("instrID") & ",document.form1.instr_id.value);'>Copy Spec</button></td>"
		Response.Write "<td>&nbsp;</td>"
		Response.Write "</tr>"
		Response.Write "</table></div>"
	End If

	Response.Write "<input type='hidden' name='flowflag' id='flowflag' value='true' />"

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