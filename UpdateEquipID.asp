<%@ language="vbscript" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="javascript">
function addItem() {
 document.form1.updateFlag.value=true;
 document.form1.submit();
}
function cancelUpdate() {
 window.close();
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Update Equipment for Instrument</title>
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
<form id="form1" name="form1" action="UpdateEquipID.asp" method="post">
<%
Dim sqlString
Dim cn
Dim rs
Dim cmd
Dim temp
Dim currentuser
Dim access

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("equipment","updateequipid",currentuser)
If access <> "none" Then

	'Define the ado connection and recordset objects.
	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")

	Response.Write "<input type='hidden' name='instrumentID' id='instrumentID' value='" & Request("instrumentID") & "' />"

	'Display the control for the user input.
	If Request("updateFlag") <> "true" Then
		Response.Write "<input type='hidden' name='updateFlag' id='updateFlag' value='false' />"
		'Get the tagname for the caption.
		sqlString = "SELECT instr_name FROM instruments WHERE instr_id=" & Request("instrumentID")
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			If Not IsNull(rs(0)) Then
				temp = rs(0)
			Else
				temp = ""
			End IF
		Else
			temp = ""
		End If
		Response.Write "<div class='center'><table style='width:100%'>"
		Response.Write "<caption><h3>Update equipment for instrument " & temp & "</h3></caption>"
		'Get the previous equipment id.
		sqlString = "SELECT equip_id FROM instruments WHERE instr_id=" & Request("instrumentID")
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			If Not IsNull(rs(0)) Then
				temp = rs(0)
			Else
				temp = 0
			End If
		Else
			temp = 0
		End If
		Response.Write "<tr>"
		Response.Write "<td class='right' style='width:30%;padding-right:5px'>Equipment:</td>"
		Response.Write "<td style='width:70%'>"
		'Load the Equipment dropdown list.
		Response.Write "<select name='equipment'>"
		sqlString = "SELECT equip_id,equip_name FROM equipment ORDER BY equip_name"
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Response.Write "<option value=''> "
			Do While Not rs.EOF
				If CLng(temp) > 0 Then
					If CLng(rs(0)) = CLng(temp) Then
						Response.Write "<option value='" & CStr(rs(0)) & "' selected>" & rs(1)
					Else
						Response.Write "<option value='" & CStr(rs(0)) & "'>" & rs(1)
					End If
				Else
					Response.Write "<option value='" & CStr(rs(0)) & "'>" & rs(1)
				End If
				rs.MoveNext
			Loop
		End If
		rs.Close
		Response.Write "</select>"
		Response.Write "<input type='hidden' name='oldEquip' id='oldEquip' value='" & temp & "' />"
		Response.Write "</td>"
		Response.Write "</tr>"
		Response.Write "</table>"

		If access = "write" Or access = "delete" Then
			Response.Write "<table style='width:100%'>"
			Response.Write "<tr>"
			Response.Write "<td style='width:50%'><input type='button' id='submit1' name='submit1' value='Save' onclick='addItem();' /></td>"
			Response.Write "<td style='width:50%'><input type='button' id='cancel' name='cancel' value='Cancel' onclick='cancelUpdate();' /></td>"
			Response.Write "</tr>"
			Response.Write "</table></div>"
		End If
	Else
		'Save the data to the instruments table.
		If Request("instrumentID") <> "" And Request("equipment") <> "" Then
			sqlString = "UPDATE instruments SET equip_id=" & Request("equipment") & _
						" WHERE instr_id=" & Request("instrumentID")
			On Error Resume Next
			Set rs = cn.Execute(sqlString)
			If Err.number <> 0 Then
				Response.Write "An error occurred equipment ID: " & Err.Description
			Else
				'Write the change to the audit trail table.
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & Request("instrumentID") & ",'instruments','equip_id','" & Request("oldEquip") & "','" & Request("equipment") & "','update')"
				Set cmd = CreateObject("adodb.command")
				cmd.ActiveConnection = cn
				cmd.CommandText = sqlString
				cmd.Execute
				Set cmd = Nothing
				Response.Write "<script language='javascript'>window.opener.location.href=window.opener.location.href;window.close();</script>"
			End If
		Else
			Response.Write "ERROR - one or more request variables do not exist!"
		End If
		
	End If

	Set rs = Nothing
	cn.Close
	Set cn = Nothing
Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If
%>
</form>
</body>
</html>