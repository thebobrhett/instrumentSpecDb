<%@ language="vbscript" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<html>
<head>
<script language="javascript">
function reloadpage() {
 document.form1.formID.value='';
 document.form1.action="selectform.asp";
 document.form1.submit();
}
function openhelp() {
 window.open("Instrument Spec Database Users Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Select Form</title>
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
<form id="form1" name="form1" action="SelectForm.asp" method="post">
<%
'*************
' Revision History
' 
' Keith Brooks - Tuesday, March 23, 2010
'   Creation
' Keith Brooks - Monday, July 30, 2012
'	Cleaned up html to remove deprecated items.
'*************

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
access = UserAccess("equipment","selectform",currentuser)
If access <> "none" Then

	'Define the ado connection and recordset objects.
	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")

	Response.Write "<input type='hidden' name='instrID' id='instrID' value='" & Request("instrID") & "' />"
	Response.Write "<input type='hidden' name='page' id='page' value='" & Request("page") & "' />"

	If Request("formID") = "" Then
		'Draw the list of specification form types.
		Response.Write "<div class='center'><table style='width:100%'>"
		Response.Write "<tr>"
		Response.Write "<td class='left top' style='width:25%'><a class='noprint' href='default.asp'>Home</a></td>"
		Response.Write "<td class='center' style='width:50%'><h3 />Select a specification form type for this instrument:</td>"
		Response.Write "<td class='right top' style='width:25%'><a class='noprint' href='' onclick='openhelp();return false;' title='Open the User Guide'>Help</a></td>"
		Response.Write "</tr>"
		Response.Write "<tr>"
		Response.Write "<td>&nbsp;</td>"
		Response.Write "<td><select name='formID' size='20'>"
		'Get the instrument function type for this instrument, then get the suggested
		'spec form for this type and highlight it in the list.
		sqlString = "SELECT instr_func_type_id FROM instruments WHERE instr_id=" & Request("instrID")
		Set rs = cn.Execute(sqlString)
		temp = 0
		If Not rs.BOF Then
			rs.MoveFirst
			temp = rs(0)
		End If
		rs.Close
		If temp > 0 Then
			sqlString = "SELECT instr_func_type_spec_form_id FROM instrument_function_types WHERE instr_func_type_id=" & temp
			Set rs = cn.Execute(sqlString)
			If Not rs.BOF Then
				rs.MoveFirst
				If Not IsNull(rs(0)) Then
					temp = rs(0)
				End If
			End If
			rs.Close
		End If
		If Request("allforms") = "" Then
			sqlString = "SELECT spec_form_id,spec_form_num,spec_form_name FROM spec_forms WHERE common_form=1 ORDER BY spec_form_num"
		Else
			sqlString = "SELECT spec_form_id,spec_form_num,spec_form_name FROM spec_forms ORDER BY spec_form_num"
		End If
		Set rs = cn.Execute(sqlString)
		If Not rs.BOF Then
			rs.MoveFirst
			Do While Not rs.EOF
				If CLng(rs(0)) = CLng(temp) Then
					Response.Write "<option value='" & rs(0) & "' selected>" & rs(1) & " - " & rs(2)
				Else
					Response.Write "<option value='" & rs(0) & "'>" & rs(1) & " - " & rs(2)
				End If
				rs.MoveNext
			Loop
		End If
		rs.Close
		Response.Write "</select></td>"
		If Request("allforms") <> "" Then
			Response.Write "<td><input type='checkbox' id='allforms' name='allforms' onclick='reloadpage();' checked />Show all forms</td>"
		Else
			Response.Write "<td><input type='checkbox' id='allforms' name='allforms' onclick='reloadpage();' />Show all forms</td>"
		End If
		Response.Write "</tr>"

		Response.Write "<tr>"
		Response.Write "<td>&nbsp;</td>"
		Response.Write "<td><input type='submit' id='submit1' name='submit1' value='Continue' /></td>"
		Response.Write "<td>&nbsp;</td>"
		Response.Write "</tr>"
		Response.Write "</table></div>"
	Else
		'Save the spec form id to the instruments table for this instrument so it
		'can be picked up by the appropriate form.
		sqlString = "UPDATE instruments SET spec_id=" & Request("formID") & " WHERE instr_id=" & Request("instrID")
		On Error Resume Next
		Set rs = cn.Execute(sqlString)
		If Err.number <> 0 Then
			Response.Write "An error occurred saving the spec form ID: " & Err.Description
		Else
			If Request("page") = "editspec" Then
				Response.Redirect "editspec.asp?instrID=" & Request("instrID") & "&page_num=1"
			ElseIf Request("page") = "printspec" Then
				Response.Redirect "printspec.asp?instrID=" & Request("instrID") & "&page_num=1"
			End If
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