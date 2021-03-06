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
function openhelp() {
 window.open("Instrument Spec Database Administrators Guide.doc","userguide");
}
<!--#include file="datepicker.js"-->
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Instrument Spec Audit Trail</title>
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
		<td class="left" style="width:20%"><a class="noprint" href="adminmenu.asp">Menu</a></td>
		<td class="center" style="width:60%"><h1 />Instrument Spec Audit Trail</td>
		<td class="right" style="width:20%"><a class="noprint" href="#" onclick="openhelp();return false;" title="Open the Admin Guide">Help</a></td>
	</tr>
</table>
<form id="form1" name="form1" action="AuditTrail.asp" method="post">
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
'access = UserAllowed(currentuser, "masterbatchentry")
access = UserAccess("equipment", "audittrail", currentuser)
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
	'Draw the criteria selection lists.
	Response.Write "<table style='width:100%;border:none'>"
	Response.Write "<tr>"
	Response.Write "<th style='width:35%'>Date Range</th>"
	Response.Write "<th style='width:20%'>Plant Area</th>"
	Response.Write "<th style='width:25%'>Instrument Tag</th>"
	Response.Write "<th style='width:20%'>Modifier</th>"
	Response.Write "</tr>"
	Response.Write "<tr>"

	Response.Write "<td>"
	Response.Write "<table style='width:100%;padding:5px'>"
	'Draw the start date.
	Response.Write "<tr>"
	response.write "<td class='right bold' style='padding:10px'>Start Date: </td>"
	If Request("start_date") <> "" Then
		Response.Write "<td class='left'><input type='text' name='start_date' size='10' value='" & Request("start_date") & "' onchange='checkDate_onchange(0)' />"
	Else
		Response.Write "<td class='left'><input type='text' name='start_date' size='10' value='' onchange='checkDate_onchange(0)' />"
	End If
	Response.Write "<a href='javascript: displayDatePicker(""start_date"");'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Response.Write "</tr>"
	'Draw the end date.
	Response.Write "<tr>"
	response.write "<td class='right bold' style='padding:10px'>End Date: </td>"
	If Request("end_date") <> "" Then
		Response.Write "<td class='left'><input type='text' name='end_date' size='10' value='" & Request("end_date") & "' onchange='checkDate_onchange(1)' />"
	Else
		Response.Write "<td class='left'><input type='text' name='end_date' size='10' value='' onchange='checkDate_onchange(1)' />"
	End If
	Response.Write "<a href='javascript: displayDatePicker(""end_date"");'><img src='../images/calendar.bmp' alt='Calendar' style='vertical-align:top'></a></td>"
	Response.Write "</tr>"
	Response.Write "</table>"
	Response.Write "</td>"

	'Load the Process Areas dropdown list.
	Response.Write "<td class='center'><select name='plantarea' onchange='doSubmit();'>"
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
	'Load the Instrument tag dropdown list.
	If Request("plantarea") <> "" Then
		Response.Write "<td><select name='tagname'>"
		sqlString = "SELECT instr_id,instr_name FROM instruments WHERE plant_area_id=" & Request("plantarea") & " ORDER BY instr_name"
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If Request("tagname") <> "" Then
					If CLng(rs(0)) = CLng(Request("tagname")) Then
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
	Else
		Response.Write "<td>&nbsp;</td>"
	End If
	'Load the modifier dropdown list.
	Response.Write "<td><select name='modifier'>"
	sqlString = "SELECT DISTINCT change_user FROM audit_trail ORDER BY change_user"
	Set rs = cn.Execute(sqlString)
	If Not rs.BOF Then
		rs.MoveFirst
		Response.Write "<option value=''> "
		Do While Not rs.EOF
			If Request("modifier") <> "" Then
				If rs(0) = Request("modifier") Then
					Response.Write "<option value='" & rs(0) & "' selected>" & rs(0)
				Else
					Response.Write "<option value='" & rs(0) & "'>" & rs(0)
				End If
			Else
				Response.Write "<option value='" & rs(0) & "'>" & rs(0)
			End If
			rs.MoveNext
		Loop
	End If
	rs.Close
	Response.Write "</select></td>"
	Response.Write "</tr>"
	Response.Write "</table>"

	Response.Write "<br />"
	Response.Write "<table style='width:100%'>"
	Response.Write "<tr class='noprint'>"
	Response.Write "<td style='width:33%'>&nbsp;</td>"
	Response.Write "<td class='center' style='width:34%'><input type='button' id='submit1' name='submit1' value='Find' style='font-size:10pt' onclick='doFind();'></td>"
	Response.Write "<td style='width:33%'>&nbsp;</td>"
	Response.Write "</tr>"
	Response.Write "</table>"

	'If any of the criteria have been selected, display the list box with the results.
	criteria = ""
	If Request("start_date") <> "" Then
		criteria = "change_date > '" & FormatMySQLDateTime(Request("start_date")) & "'"
	End If
	If Request("end_date") <> "" Then
		If criteria = "" Then
			criteria = "change_date < '" & FormatMySQLDateTime(DateAdd("d",1.0,Request("end_date"))) & "'"
		Else
			criteria = criteria & " AND change_date < '" & FormatMySQLDateTime(DateAdd("d",1.0,Request("end_date"))) & "'"
		End If
	End If
	If Request("plantarea") <> "" Then
		If criteria = "" Then
			criteria = "plant_area_id=" & Request("plantarea")
		Else
			criteria = criteria & " AND plant_area_id=" & Request("plantarea")
		End If
	End If
	If Request("tagname") <> "" Then
		If criteria = "" Then
			criteria = "change_instr_id=" & Request("tagname")
		Else
			criteria = criteria & " AND change_instr_id=" & Request("tagname")
		End If
	End If
	If Request("modifier") <> "" Then
		If criteria = "" Then
			criteria = "change_user='" & Request("modifier") & "'"
		Else
			criteria = criteria & " AND change_user='" & Request("modifier") & "'"
		End If
	End If
	If Request("flowflag") = "true" And criteria <> "" Then
		sqlString = "SELECT change_date,change_user,instr_name,CONCAT(change_table,'.',change_field),old_value,new_value,change_type " & _
					"FROM audit_trail LEFT JOIN instruments ON audit_trail.change_instr_id=instruments.instr_id WHERE " & criteria & " ORDER BY audit_trail_id"
		Set rs = cn.Execute(sqlString)
		Response.Write "<table style='width:100%'>"
		Response.Write "<tr>"
		Response.Write "<th class='mediumth'>Timestamp</th>"
		Response.Write "<th class='mediumth'>Modifier</th>"
		Response.Write "<th class='mediumth'>Instrument</th>"
		Response.Write "<th class='mediumth'>Table.Field</th>"
		Response.Write "<th class='mediumth'>Old Value</th>"
		Response.Write "<th class='mediumth'>New Value</th>"
		Response.Write "<th class='mediumth'>Change Type</th>"
		Response.Write "</tr>"
		If Not rs.BOF Then
			rs.MoveFirst
			Do While Not rs.EOF
				Response.Write "<tr>"
				If Not IsNull(rs(0)) Then
					Response.Write "<td class='mediumtd'>" & rs(0) & "</td>"
				Else
					Response.Write "<td class='mediumtd'>&nbsp;</td>"
				End If
				If Not IsNull(rs(1)) Then
					Response.Write "<td class='mediumtd'>" & rs(1) & "</td>"
				Else
					Response.Write "<td class='mediumtd'>&nbsp;</td>"
				End If
				If Not IsNull(rs(2)) Then
					Response.Write "<td class='mediumtd'>" & rs(2) & "</td>"
				Else
					Response.Write "<td class='mediumtd'>&nbsp;</td>"
				End If
				If Not IsNull(rs(3)) Then
					Response.Write "<td class='mediumtd'>" & rs(3) & "</td>"
				Else
					Response.Write "<td class='mediumtd'>&nbsp;</td>"
				End If
				If Not IsNull(rs(4)) Then
					Response.Write "<td class='mediumtd'>" & rs(4) & "</td>"
				Else
					Response.Write "<td class='mediumtd'>&nbsp;</td>"
				End If
				If Not IsNull(rs(5)) And rs(5) <> " " And rs(5) <> "" Then
					Response.Write "<td class='mediumtd'>" & rs(5) & "</td>"
				Else
					Response.Write "<td class='mediumtd'>&nbsp;</td>"
				End If
				If Not IsNull(rs(6)) Then
					Response.Write "<td class='mediumtd'>" & rs(6) & "</td>"
				Else
					Response.Write "<td class='mediumtd'>&nbsp;</td>"
				End If
				Response.Write "</tr>"
				rs.MoveNext
			Loop
		End If
		rs.Close
		Response.Write "</table>"
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
<script language="VBScript">
<!--
Function checkDate_onchange(index)
	Dim strDate
	On Error Resume Next
	If index = 0 Then
 		strDate = document.form1.start_date.value
 		strDate = FormatDateTime(strDate,vbShortDate)
	ElseIf index = 1 Then
 		strDate = document.form1.end_date.value
 		strDate = FormatDateTime(strDate,vbShortDate)
 	End If
	If Err <> 0 Then
		MsgBox "Invalid date format entered: " & strDate
	End If
End Function
//-->
</script>
</html>