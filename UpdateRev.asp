<%@ language="vbscript" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="javascript">
function addItem() {
 if (document.form1.rev_number.value == '') {
  alert('You must enter a revision number');
 } else {
  if (document.form1.rev_date.value == '') {
   alert('You must enter a revision date');
  } else {
   document.form1.submit();
  }
 }
}
function cancelUpdate() {
 window.close();
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Update Spec Revision</title>
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
<form id="form1" name="form1" action="UpdateRev.asp" method="post">
<%
Dim sqlString
Dim cn
Dim rs
Dim temp
Dim currentuser
Dim access

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("equipment","updaterev",currentuser)
If access <> "none" Then

	'Define the ado connection and recordset objects.
	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")

	Response.Write "<input type='hidden' name='instrID' id='instrID' value='" & Request("instrID") & "' />"

	'If the rev number hasn't been specified yet, the form has just been opened,
	'so display the controls for the user input.
	If Request("rev_number") = "" Then
		'Get the tagname for the caption.
		sqlString = "SELECT instr_name FROM instruments WHERE instr_id=" & Request("instrID")
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
		Response.Write "<caption><h3>Enter Spec Revision Information for instrument " & temp & "</h3></caption>"
		'Get the previous revision number.
		sqlString = "SELECT rev_number FROM revisions WHERE instr_id=" & Request("instrID") & " ORDER BY rev_date DESC, rev_number DESC"
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			If Not IsNull(rs(0)) Then
				temp = rs(0)
			Else
				temp = ""
			End If
		Else
			temp = ""
		End If
		If temp = "" Then
			temp = 0
		ElseIf IsNumeric(temp) Then
			temp = CInt(temp) + 1
		Else
			temp = 0
		End If
		Response.Write "<tr>"
		Response.Write "<td class='right' style='width:30%;padding-right:5px'>Rev. Number:</td>"
		Response.Write "<td class='left' style='width:70%'><input type='text' id='rev_number' name='rev_number' maxlength='3' size='5' value='" & temp & "' /></td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td class='right' style='padding-right:5px'>Rev. Date:</td>"
		Response.Write "<td class='left'><input type='text' id='rev_date' name='rev_date' size='10' value='" & FormatDateTime(Now,2) & "' /></td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td class='right' style='padding-right:5px'>Description:</td>"
		Response.Write "<td class='left'><textarea id='rev_desc' name='rev_desc' cols='50' rows='3'></textarea></td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td class='right' style='padding-right:5px'>Short Description:</td>"
		Response.Write "<td class='left'><input type='text' id='rev_short_desc' name='rev_short_desc' maxlength='35' size='40' /></td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td class='right' style='padding-right:5px'>User Initials:</td>"
		Response.Write "<td class='left'><input type='text' id='rev_user_initials' name='rev_user_initials' maxlength='3' size='3' /></td>"
		Response.Write "</tr>"
		Response.Write "</table>"

		Response.Write "<table style='width:100%'>"
		Response.Write "<tr>"
		Response.Write "<td style='width:50%'><input type='button' id='submit1' name='submit1' value='Save' onclick='addItem();' /></td>"
		Response.Write "<td style='width:50%'><input type='button' id='cancel' name='cancel' value='Cancel' onclick='cancelUpdate();' /></td>"
		Response.Write "</tr>"
		Response.Write "</table></div>"
	Else
		'Get the current user.
		temp = Request.ServerVariables("LOGON_USER")
		temp = Right(temp,Len(temp)-InStr(temp,"\"))

		'Save the data to the revisions table.
		sqlString = "INSERT INTO revisions " & _
					"(instr_id,rev_number,rev_date,rev_user,rev_desc,rev_short_desc,rev_user_initials) " & _
					"VALUES (" & Request("instrID") & ",'" & Request("rev_number") & "','" & _
					FormatMySQLDateTime(Request("rev_date")) & "','" & temp & "','" & _
					Request("rev_desc") & "','" & Request("rev_short_desc") & "','" & _
					Request("rev_user_initials") & "')"
		On Error Resume Next
		Set rs = cn.Execute(sqlString)
		If Err.number <> 0 Then
			Response.Write "An error occurred saving the spec revision info: " & Err.Description
		Else
			Response.Write "<script language='javascript'>window.opener.location.href=window.opener.location.href;window.close();</script>"
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